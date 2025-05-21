import time
import win32com.client
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def main():
    # Configurar o driver do Selenium
    options = webdriver.ChromeOptions()
    options.headless = False
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Acessar a página da VTEX
    driver.get('https://adcosprofessional.myvtex.com/admin/b2b-organizations/organizations/#/requests')
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'email')))

    # Login no VTEX
    driver.find_element(By.ID, 'email').send_keys('eduardo.klitzke@adcos.com.br')
    driver.find_element(By.CSS_SELECTOR, '[data-testid="email-form-continue"]').click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'password')))
    driver.find_element(By.NAME, 'password').send_keys('Breakmen451.')
    driver.find_element(By.ID, 'chooseprovider_signinbtn').click()

    # Esperar o iframe carregar
    iframe = WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'iframe[title="IO iframe container"]'))
)
    driver.switch_to.frame(iframe)

    # Aplicar filtro e selecionar 100 itens por página
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.c-action-primary')))
    botao_status = driver.find_element(By.CSS_SELECTOR, 'button.c-action-primary')
    botao_status.click()

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.fixed.absolute-ns.w-100.w-auto-ns.z-999.ba.bw1.b--muted-4.bg-base.left-0.br2.o-100')))
    driver.find_element(By.CSS_SELECTOR, 'input[name="status-checkbox-group"][value="approved"]').click()
    driver.find_element(By.CSS_SELECTOR, 'input[name="status-checkbox-group"][value="declined"]').click()

    driver.find_element(By.XPATH, '//button[div[contains(@class, "vtex-button__label") and contains(text(), "Apply")]]').click()

    dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'select.o-0.absolute.top-0.left-0.h-100.w-100.bottom-0.t-body.pointer')))
    dropdown.click()
    dropdown.find_element(By.XPATH, "//option[text()='100']").click()

    time.sleep(5)  # Espera carregamento inicial

    emails = []
    datas = []

    # Função para coletar emails e datas na página atual
    def coletar_emails_datas():
        elementos = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#render-admin\\.app\\.b2b-organizations\\.organizations > div > div.admin-ui-c-ervJfA > div > div > div.flex.flex-column > div:nth-child(1) > div.vh-100.w-100.dt > div > div:nth-child(1) > div > div:nth-child(2) > div:nth-child(2) > div'))
        )
        todos_dados = [element.text for element in elementos]
        dados_unidos = ''.join(todos_dados)
        dados_separados = dados_unidos.split('\n')

        for i in range(len(dados_separados)):
            if i % 3 == 0 and 'Aprovado' not in dados_separados[i+1]:
                emails.append(dados_separados[i])
                datas.append(dados_separados[i+2])

    # Coletar da primeira página
    coletar_emails_datas()

    # Passar para a segunda página e coletar novamente
    try:
        button = driver.find_element(By.CSS_SELECTOR, 'button.vtex-button.icon-button.bg-action-secondary')
        button.click()
        
        time.sleep(5)  # Aguardar carregamento da segunda página
        coletar_emails_datas()
    except:
        print("Não foi possível acessar a segunda página.")

    resultado = list(zip(emails, datas))

    # Abrir Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    planilha = excel.ActiveWorkbook.ActiveSheet

    linha_inicial = 16
    ultimo_email = None

    # Encontrar último email não aprovado
    while True:
        email_na_linha = planilha.Range(f'B{linha_inicial}').Value
        status_na_linha = planilha.Range(f'R{linha_inicial}').Value
        if status_na_linha != "Sim":
            ultimo_email = email_na_linha
            break
        linha_inicial += 1

    novos_resultados = []
    encontrou_ultimo_email = False
    for email, data in resultado:
        if email == ultimo_email:
            encontrou_ultimo_email = True
            break
        if not encontrou_ultimo_email:
            novos_resultados.append((email, data))

    linha_inicial = 16
    Id = 0
    planilha.Columns("H:H").Hidden = False
    planilha.Columns("P:P").Hidden = False
    planilha.Columns("Q:Q").Hidden = False

    for email, data in novos_resultados:
        planilha.Rows(linha_inicial).Insert()
        try:
            data_formatada = datetime.strptime(data, '%m/%d/%Y').strftime('%d/%m/%Y')
        except ValueError:
            data_formatada = data
        planilha.Range(f'A{linha_inicial}').Value = Id
        planilha.Range(f'B{linha_inicial}').Value = email
        planilha.Range(f'C{linha_inicial}').Value = data_formatada
        planilha.Range(f'H{linha_inicial}').FormulaLocal = f'=ESQUERDA(G{linha_inicial}; 3)'
        planilha.Range(f'P{linha_inicial}').FormulaLocal = f'=SE(O{linha_inicial}="Aprovado";"Aprovado";SE(O{linha_inicial}<>"";"Comunicar";""))'
        planilha.Range(f'Q{linha_inicial}').FormulaLocal = f'=H{linha_inicial}&P{linha_inicial}'
        linha_inicial += 1

    planilha.Columns("H:H").Hidden = True
    planilha.Columns("P:P").Hidden = True
    planilha.Columns("Q:Q").Hidden = True
    excel.ActiveWorkbook.Save()
    driver.quit()

if __name__ == '__main__':
    main()
