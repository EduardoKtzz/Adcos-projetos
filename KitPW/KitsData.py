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


def alterar_data(log, log_area):
    # ABRE EXCEL - Acessa a planilha aberta no Excel
    excel = win32com.client.Dispatch("Excel.Application")
    abaplanilha = None

    # CAÇA A PLANILHA - Verifica todas planilhas para achar certa
    for abaplanilha in excel.Workbooks:
        try:
            abaplanilha = abaplanilha.Sheets("ALTERAR_DATA")
            log(f"Planilha encontrada na pasta de trabalho", log_area)
            break
        except Exception as e:
            log(f"Planilha não encontrada na pasta de trabalho", log_area)

    # NAO ACHAR PLANILHA - Se a planilha não foi encontrada, encerra o script
    if not abaplanilha:
        log("🚨 ERRO: A planilha não foi encontrada em nenhuma pasta de trabalho aberta.", log_area)
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

    # PREENCHER KITS - Coleta os dados da planilha
    for linha in range(2, ultima_linha + 1):  

        # DATA - Pega e formata as datas
        data_inicial = abaplanilha.Cells(linha, 2).Value # COLUNA B
        data_final = abaplanilha.Cells(linha, 3).Value  # COLUNA C

        # Formata as datas
        if isinstance(data_inicial, datetime) and isinstance(data_final, datetime):
            data_inicial = data_inicial.strftime("%d/%m/%Y")
            data_final = data_final.strftime("%d/%m/%Y")

        # Atualizar datas dos kits
        url_kit = abaplanilha.Cells(linha, 1).Value  # COLUNA A (URL do kit)

        if not url_kit:
            log(f"⚠️ Linha {linha}: URL do kit ausente. Pulando...", log_area)
            continue

        try:
            url_kit = int(url_kit)
        except ValueError:
            log(f"⚠️ Linha {linha}: URL do kit inválida. Pulando...",log_area)
            continue
        
        # Acessa o formulário do kit
        url_formulario = f"https://pw.adcos.com.br/backend/kits/form/{url_kit}/"
        driver.get(url_formulario)

        log(f"Acessando: {url_kit}", log_area)

        #DATA INICIAL
        campo_data_inicial = driver.find_element(By.NAME, "data_inicial")
        pyperclip.copy(data_inicial)
        campo_data_inicial.click()
        campo_data_inicial.send_keys(Keys.CONTROL, 'v')

        campo_data_final = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "data_final"))
        )

        # Exibe as datas para conferência
        log(f"Data Final: {data_inicial}", log_area)
        log(f"Data Final: {data_final}", log_area)
    
        # Atualiza o campo de data final
        pyperclip.copy(data_final)
        campo_data_final.click()
        campo_data_final.send_keys(Keys.CONTROL, 'v')

        # Enviar formulário
        driver.find_element(By.ID, "bt-salvar").click()
        time.sleep(2)  # Pequeno delay para evitar problemas

    print("✅ Datas dos kits atualizadas com sucesso!")
    driver.quit()
