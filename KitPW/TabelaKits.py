# Importações do projeto
import sys
import time
import win32com.client
import pyperclip
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ABRE EXCEL - Acessa a planilha aberta no Excel
excel = win32com.client.Dispatch("Excel.Application")
abaplanilha = None

# CAÇA A PLANILHA - Verifica todas planilhas para achar certa de KITS
for abaplanilha in excel.Workbooks:
    try:
        abaplanilha = abaplanilha.Sheets("INSERIR_TABELA")
        print(f"Planilha 'Cadastro_KITS' encontrada na pasta de trabalho '{abaplanilha.Name}'")
        break
    except Exception as e:
        print(f"Planilha 'Cadastro_KITS' não encontrada na pasta de trabalho '{abaplanilha.Name}'")

# NAO ACHAR PLANILHA - Se a planilha não foi encontrada, encerra o script
if not abaplanilha:
    print("🚨 ERRO: A planilha 'Cadastro de Kits' não foi encontrada em nenhuma pasta de trabalho aberta.")
    sys.exit("Encerrando o robô...")

# TOTAL DE LINHA - Descobre o número total de linhas preenchidas na planilha
ultima_linha = abaplanilha.Cells(abaplanilha.Rows.Count, 1).End(3).Row  

# SELENIUM - Configurar o driver do Selenium
options = webdriver.ChromeOptions()
options.headless = False
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# ACESSAR SITE - Acessar a página de login do PW
driver.get('https://pw.adcos.com.br/backend/kits/')
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'usuario')))

# LOGIN - Fazer login no HPW
driver.find_element(By.NAME, 'usuario').send_keys('Eduardo.Klitzke')
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'senha')))
driver.find_element(By.NAME, 'senha').send_keys('C8vWr')
driver.find_element(By.CLASS_NAME, "btn").click()
time.sleep(2)

# NOME E PRODUTO - Cria estrutura para armazenar os kits
kits_dict = {}
kit_codigo = 0

# Função auxiliar para validar se a célula está preenchida
def obter_valor_celula(celula):
    return str(int(celula)) if celula is not None else None


# PREENCHER FORMULÁRIO - Automatiza o preenchimento
for linha in range(2, ultima_linha + 1):
    nome_kit = abaplanilha.Cells(linha, 4).Value  # COLUNA D (Nome do kit)
    produto_codigo = obter_valor_celula(abaplanilha.Cells(linha, 6).Value)  # COLUNA F (Código do produto)

    # Código específico do formulário (COLUNA I)
    codigo_formulario = obter_valor_celula(abaplanilha.Cells(linha, 9).Value)

    # Valida se o código do formulário está preenchido
    if not codigo_formulario:
        print(f"⚠️ Linha {linha}: Código do formulário ausente. Pulando...")
        continue

    # Preços
    preco_padrao = abaplanilha.Cells(linha, 7).Value or 0  # COLUNA G
    preco_desconto = abaplanilha.Cells(linha, 8).Value or 0  # COLUNA H

    # Inicializa o kit no dicionário se não existir
    if nome_kit not in kits_dict:
        kits_dict[nome_kit] = {"produtos": [], "total_preco_padrao": 0, "total_preco_desconto": 0}

    # Adicionar produto ao kit
    kits_dict[nome_kit]["produtos"].append({
        "produto": produto_codigo,
        "preco_padrao": preco_padrao,
        "preco_desconto": preco_desconto,
        "codigo_formulario": codigo_formulario
    })

    # Atualiza os totais
    kits_dict[nome_kit]["total_preco_padrao"] += preco_padrao
    kits_dict[nome_kit]["total_preco_desconto"] += preco_desconto

# Lista para armazenar os links já processados
links_processados = set()

# PREENCHER FORMULÁRIO - Automatiza o preenchimento com base no código do formulário
for nome_kit, dados_kit in kits_dict.items():
    produtos = dados_kit["produtos"]
    total_preco_padrao = dados_kit["total_preco_padrao"]
    total_preco_desconto = dados_kit["total_preco_desconto"]

    for produto in produtos:
        codigo_formulario = produto.get("codigo_formulario")

        # Ignorar se o código do formulário estiver ausente
        if not codigo_formulario:
            print(f"⚠️ Produto sem código de formulário em {nome_kit}. Pulando...")
            continue

        # Acessa o link com o código correto
        url_formulario = f"https://pw.adcos.com.br/backend/kits/form/{codigo_formulario}/"

        # Verificar se o link já foi processado
        if url_formulario in links_processados:
            print(f"🔁 Link já processado para o kit {nome_kit}. Pulando...")
            continue  # Pular a linha caso o link já tenha sido processado

        # Se o link ainda não foi processado, adicione ao conjunto de links processados
        links_processados.add(url_formulario)

        driver.get(url_formulario)

        print(f"Acessando: {url_formulario}")

        # TABELA DE PRECOS
        tabela_precos = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@rel='box-tabela-de-prea-os']/span[text()='Tabela de preços']"))
        )
        tabela_precos.click()

        # Esperar o botão ficar visível e clicável
        botao_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "dropdown-toggle"))
        )

        # Clicar no botão para abrir o dropdown
        botao_dropdown.click()
        time.sleep(2)

        # Aguarde até que o elemento esteja visível
        opcao = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='text' and text()='A72 - CLIENTE FINAL COM PRESCRICAO']"))
        )
        opcao.click()  # Agora você pode interagir com o elemento

        total_preco_desconto = round(total_preco_desconto, 2)
        total_preco_desconto = f"{total_preco_desconto:.2f}".replace('.', ',')

        # Agora, você pode enviar esses valores formatados para o formulário
        campo_preco_padrao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_vista"))
        )
        campo_preco_padrao.clear()
        campo_preco_padrao.send_keys(total_preco_desconto)

        campo_preco_padrao1 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_prazo"))
        )

        campo_preco_padrao1.click()
        campo_preco_padrao1.clear()
        campo_preco_padrao1.send_keys(total_preco_desconto)
        campo_preco_padrao1.send_keys(Keys.RETURN)

        driver.refresh()

        tabela_precos = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@rel='box-tabela-de-prea-os']/span[text()='Tabela de preços']"))
        )
        tabela_precos.click()

        # Esperar o botão ficar visível e clicável
        botao_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "dropdown-toggle"))
        )

        # Clicar no botão para abrir o dropdown
        botao_dropdown.click()

        time.sleep(2)
        # Aguarde até que o elemento esteja visível
        opcao_2 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='text' and text()='A660 - 10% OFF REF A19_NOVOS CLIENTES']"))
        )
        opcao_2.click()

        # Agora, você pode enviar esses valores formatados para o formulário
        campo_preco_desconto = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_vista"))
        )
        campo_preco_desconto.clear()
        campo_preco_desconto.send_keys(total_preco_desconto)

        campo_preco_desconto1 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_prazo"))
        )

        campo_preco_desconto1.click()
        campo_preco_desconto1.clear()
        campo_preco_desconto1.send_keys(total_preco_desconto)
        campo_preco_desconto1.send_keys(Keys.RETURN)

        time.sleep(2)

print("✅ Formulários preenchidos com sucesso!")
driver.quit()