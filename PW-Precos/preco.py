# Importações do projeto
import sys
import time
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# Acessa a planilha aberta no Excel
excel = win32com.client.Dispatch("Excel.Application")
ws = None

# Itera sobre todas as pastas de trabalho abertas
for wb in excel.Workbooks:
    # Verifica se a planilha "itens" existe na pasta de trabalho atual
    try:
        ws = wb.Sheets("itens")  # Tenta acessar a aba pelo nome
        print(f"Planilha 'Itens' encontrada na pasta de trabalho '{wb.Name}'")
        break  # Sai do loop se encontrar a planilha desejada
    except Exception as e:
        print(f"Planilha 'Itens' não encontrada na pasta de trabalho '{wb.Name}'")

    # Se a planilha não foi encontrada, encerra o script
if not ws:
    print("🚨 ERRO: A planilha 'MATERIAL' não foi encontrada em nenhuma pasta de trabalho aberta.")
    sys.exit("Encerrando o robô...")

# Pega os códigos da coluna A,B,C e D parando na primeira célula vazia
dados = []
linha = 2  
while True:

    # Posição das variaveis na lista
    material = ws.Cells(linha, 1).Value
    valor = ws.Cells(linha, 2).Value
    ncm = ws.Cells(linha, 3).Value
    hierarquia = str(ws.Cells(linha, 4).Value).zfill(5)

    # Para quando encontrar uma linha vazia
    if material is None: 
        break

    # Substitui a vírgula por ponto antes de converter para float
    valor_formatado = "{:.2f}".format(float(str(valor).replace(',', '.')))

    # Adiciona os dados na lista como um dicionário
    dados.append({ 
        "linha": linha,
        "material": str(int(material)),
        "valor": valor_formatado, 
        "ncm": str(ncm),
        "hierarquia": hierarquia
    })

    # Vai para a próxima linha
    linha += 1 

# SELENIUM - Configurar o driver do Selenium
options = webdriver.ChromeOptions()
options.headless = False
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# ACESSAR SITE - Acessar a página de login do PW
driver.get('https://pw.adcos.com.br/backend/kits/')
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'usuario')))

# LOGIN - Fazer login no PW
driver.find_element(By.NAME, 'usuario').send_keys('Eduardo.Klitzke')
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'senha')))
driver.find_element(By.NAME, 'senha').send_keys('C8vWr')
driver.find_element(By.CLASS_NAME, "btn").click()
time.sleep(2)

# Percorre todos os materiais da planilha e faz as alterações no formulario
for item in dados:
    linha_planilha = item["linha"]  # Pegamos a linha correspondente ao MATERIAL
    codigo = item["material"]
    valor = item["valor"]
    ncm = item["ncm"]
    hierarquia = item["hierarquia"]

    # Acesse a URL específica do produto
    driver.get(f"https://pw.adcos.com.br/backend/produtos/form/{codigo}/")
    time.sleep(3)

    # Insere o preço
    campo_valor = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "valor"))
    )

    #Rola até o campo, limpa o campo e insere o valor da planilha
    driver.execute_script("arguments[0].scrollIntoView();", campo_valor)
    campo_valor.click()
    campo_valor.clear()
    campo_valor.send_keys(valor)

    # Insere o NCM
    campo_ncm = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "ncm"))
    )
    campo_ncm.clear()
    campo_ncm.send_keys(ncm)

    # Insere hierarquia
    campo_hierarquia = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "hierarquia"))
    )

    campo_hierarquia.clear()
    campo_hierarquia.send_keys(hierarquia)

    time.sleep(50)

    # Clica no botão de salvar
    botao_salvar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "bt-salvar"))
    )
    botao_salvar.click()

    # Marca "X" na coluna E (Status) na linha correta do MATERIAL
    try:
        ws.Cells(linha_planilha, 5).Value = "X"
        print(f"✅ Produto {codigo} atualizado com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao atualizar o produto {codigo}: {e}")
        ws.Cells(linha_planilha, 5).Value = "Erro"

# Fechar o navegador
driver.quit()
