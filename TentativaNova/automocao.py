import time
import win32com.client
from datetime import datetime
import pythoncom
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

# Configurando uma função para esperar alguma coisa carregar
def esperar(driver, by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def rodar_automacao(email, senha):

    pythoncom.CoInitialize()

    # Configurar o driver do Selenium
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-infobars")  # Remove barra de info do Chrome
    options.headless = False
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Acessar a página da VTEX
    driver.get('https://adcosprofessional.myvtex.com/admin/b2b-organizations/organizations/#/requests')

    # Login no VTEX
    # Email
    esperar(driver, By.ID, 'email').send_keys(email)
    driver.find_element(By.CSS_SELECTOR, '[data-testid="email-form-continue"]').click()

    # Senha
    try:
        esperar(driver, By.NAME, 'password', timeout=10).send_keys(senha)
    except TimeoutException:
        driver.quit()
        raise Exception("Email externo detectado. Use um email corporativo.")
        
    driver.find_element(By.ID, 'chooseprovider_signinbtn').click()

    # Verifica se o login foi feito corretamente
    try:
        iframe = esperar(driver, By.CSS_SELECTOR, 'iframe[title="IO iframe container"]')
        driver.switch_to.frame(iframe)
    except TimeoutException:
        driver.quit()
        raise Exception("Falha no login. Verifique seu email e senha")

    # Clica no botão de filtros
    esperar(driver, By.CSS_SELECTOR, 'button.c-action-primary').click()
    
    # Espera o filtro carregar e depois clica para retirar cadastros aprovados
    esperar(driver, By.CSS_SELECTOR, '.fixed.absolute-ns.w-100.w-auto-ns.z-999.ba.bw1.b--muted-4.bg-base.left-0.br2.o-100')
    driver.find_element(By.CSS_SELECTOR, 'input[name="status-checkbox-group"][value="approved"]').click()
  
    # Clica no botão de aplicar
    esperar(driver, By.XPATH, '//button[div[contains(@class, "vtex-button__label") and contains(text(), "Apply")]]').click()

    # Espera a opção de ampliar para 100 solicitações carregar e clica nela
    esperar(driver, By.CSS_SELECTOR, 'select.o-0.absolute.top-0.left-0.h-100.w-100.bottom-0.t-body.pointer').click()
    esperar(driver, By.XPATH, "//option[text()='100']").click()

    # Esperar carregamento da nova tabela
    time.sleep(2)

    # Configurando o excel para o melhor desempenho
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.ScreenUpdating = False
    excel.Calculation = -4135  # xlCalculationManual
    excel.EnableEvents = False

    # Acessando a planilha no excel
    ws_dados = None
    encontrou = False
    for wb in excel.Workbooks:
        for s in wb.Sheets:
            if s.Name.strip().upper() == "CLIENTES VTEX":
                ws_dados = s
                encontrou = True
                break
        if encontrou:
            break

    # Se não encontrar, lança um erro visível
    if ws_dados is None:
        raise Exception("❌ A aba 'CLIENTES VTEX' não está aberta no Excel. Por favor, abra a planilha antes de executar.")

    # Definindo as variaveis
    linha_inicial = 16
    linha_insercao = 16
    ultimo_email = None

    # Encontrar último email não aprovado
    while True:
        email_na_linha = ws_dados.Range(f'B{linha_inicial}').Value
        status_na_linha = ws_dados.Range(f'R{linha_inicial}').Value

        # Se a célula estiver vazia, finaliza
        if not email_na_linha:
            break

        if status_na_linha is None or status_na_linha in ("", "Não"):
            ultimo_email = email_na_linha
            break

        linha_inicial += 1

    # Criando arrays
    emails_novos = []
    datas_novas = []
    encontrou_ultimo = False

    #Função para coletar todos os emails da pagina
    def coletar_emails_datas():
        try:
            elementos = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#render-admin\\.app\\.b2b-organizations\\.organizations > div > div.admin-ui-c-ervJfA > div > div > div.flex.flex-column > div:nth-child(1) > div.vh-100.w-100.dt > div > div:nth-child(1) > div > div:nth-child(2) > div:nth-child(2) > div')))
            dados = "\n".join(el.text for el in elementos).split('\n')

            #Agrupando de 3 em 3
            for i in range(0, len(dados), 3):
                try:
                    email = dados[i]
                    status = dados[i + 1]
                    data = dados[i + 2]

                    if email == ultimo_email:
                        return True

                    emails_novos.append(email)
                    datas_novas.append(data)

                except IndexError:
                    # Caso a linha esteja incompleta, apenas ignora
                    continue

        except TimeoutException:
            raise Exception("❌ Não foi possível localizar os dados na tabela.")
        return False

    encontrou_ultimo = coletar_emails_datas()

    if not encontrou_ultimo:

        # Passar para a segunda página e coletar novamente
        try:
            driver.find_element(By.CSS_SELECTOR, 'button.vtex-button.icon-button.bg-action-secondary').click()
            time.sleep(3)
            coletar_emails_datas()
        except:
            print("Não foi possível acessar a segunda página.")

    # Faz a inserção de email tudo de uma vez
    qtd = len(emails_novos)

    if qtd > 0:
        ws_dados.Columns("H:H").Hidden = False
        ws_dados.Columns("P:P").Hidden = False
        ws_dados.Columns("Q:Q").Hidden = False

        ws_dados.Rows(f"{linha_insercao}:{linha_insercao + qtd - 1}").Insert()
        ws_dados.Range(f"A{linha_insercao}:A{linha_insercao + qtd - 1}").Value = [[0]] * qtd
        ws_dados.Range(f"B{linha_insercao}:B{linha_insercao + qtd - 1}").Value = [[e] for e in emails_novos]

        datas_formatadas = []
        for data in datas_novas:
            try:
                data_formatada = datetime.strptime(data, '%m/%d/%Y').strftime('%d/%m/%Y')
            except ValueError:
                data_formatada = data
            datas_formatadas.append([data_formatada])

        ws_dados.Range(f"C{linha_insercao}:C{linha_insercao + qtd - 1}").Value = datas_formatadas

        for i in range(linha_insercao, linha_insercao + qtd):
            ws_dados.Range(f'H{i}').FormulaLocal = f'=ESQUERDA(G{i}; 3)'
            ws_dados.Range(f'P{i}').FormulaLocal = f'=SE(O{i}="Aprovado";"Aprovado";SE(O{i}<>"";"Comunicar";""))'
            ws_dados.Range(f'Q{i}').FormulaLocal = f'=H{i}&P{i}'

        ws_dados.Columns("H:H").Hidden = True
        ws_dados.Columns("P:P").Hidden = True
        ws_dados.Columns("Q:Q").Hidden = True

    excel.ScreenUpdating = True
    excel.Calculation = -4105  # xlCalculationAutomatic
    excel.EnableEvents = True

    excel.ActiveWorkbook.Save()

    pythoncom.CoUninitialize()
    
    driver.quit()

