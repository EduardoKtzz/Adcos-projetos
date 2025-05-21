# Importações do projeto
import sys
import time
import win32com.client
import pyperclip
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def preencherKits(log, log_area, kit_codigo):
    # ABRE EXCEL - Acessa a planilha aberta no Excel
    excel = win32com.client.Dispatch("Excel.Application")
    abaplanilha = None

    # CAÇA A PLANILHA - Verifica todas planilhas para achar a certa de KITS
    for workbook in excel.Workbooks:
        try:
            # Tenta acessar a aba "CADASTRO KIT"
            abaplanilha = workbook.Sheets("CADASTRO KIT")
            log(f"Planilha de cadastro de kits encontrada em '{abaplanilha.Name}'", log_area)
        except Exception:
            # Captura o erro e mostra no log para facilitar o diagnóstico
            log("Erro ao procurar a aba em: " + workbook.Name, log_area)
        

    # NAO ACHAR PLANILHA - Se a planilha não foi encontrada, encerra o script
    if not abaplanilha:
        log("A planilha não foi encontrada em nenhuma pasta de trabalho aberta.", log_area)
        sys.exit("Encerrando o robô...")

  
    #TOTAL DE LINHAS
    ultima_linha = abaplanilha.Cells(abaplanilha.Rows.Count, 1).End(3).Row
      

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

    # ESTRUTURA DE KITS E NUMERO DO KIT - Cria estrutura para armazenar os kits
    kits_dict = {}
    
    # PEGAR DADOS DA PLANILHA - Coleta os dados da planilha
    for linha in range(2, ultima_linha + 1):
        # Nome do kit e o produto  
        nome_kit = abaplanilha.Cells(linha, 4).Value  # COLUNA D
        produto_codigo = str(int(abaplanilha.Cells(linha, 6).Value)) # COLUNA F

        # Preços
        preco_padrao = abaplanilha.Cells(linha, 7).Value  # COLUNA G
        preco_desconto = abaplanilha.Cells(linha, 8).Value  # COLUNA H

        # Datas
        data_inicial = abaplanilha.Cells(linha, 1).Value  # COLUNA A
        data_final = abaplanilha.Cells(linha, 2).Value  # COLUNA B

        # Formata as datas
        if isinstance(data_inicial, datetime) and isinstance(data_final, datetime):
            data_inicial = data_inicial.strftime("%d/%m/%Y")
            data_final = data_final.strftime("%d/%m/%Y")

        # CRIAR KITS - Cria um kit para cada produto individualmente
        if nome_kit not in kits_dict:
            kits_dict[nome_kit] = {"produtos": [], "total_preco_padrao": 0, "total_preco_desconto": 0, "linha":linha}

        # Adicionar cada produto ao kit, incluindo o cálculo de preço
        kits_dict[nome_kit]["produtos"].append({
            "produto": produto_codigo,
            "data_inicial": data_inicial,
            "data_final": data_final,
            "preco_padrao": preco_padrao,
            "preco_desconto": preco_desconto
        })

        # Soma os valores e adciona no kit
        kits_dict[nome_kit]["total_preco_padrao"] += preco_padrao
        kits_dict[nome_kit]["total_preco_desconto"] += preco_desconto

    # PREENCHER FORMULÁRIO - Automatiza o preenchimento
    for nome_kit, dados_kit in kits_dict.items():
        linha_kit = dados_kit["linha"]
        produtos = dados_kit["produtos"]
        total_preco_padrao = dados_kit["total_preco_padrao"]
        total_preco_desconto = dados_kit["total_preco_desconto"]

        nome_personalizado = f"{nome_kit} ({'-'.join(produto['produto'] for produto in produtos)})"
        kit_codigo_formatado = f"KIT {str(kit_codigo).zfill(2)}"
        kit_codigo += 1
        hora_inicial = "00:01"
        hora_final = "23:59"
        quantidade = 1

        #Print para ver em qual kit está
        log(f"Preenchendo formulario do: {kit_codigo_formatado} - {nome_personalizado}\n", log_area)

        # NOVO KIT - Abre o link para criar um novo kit
        driver.get("https://pw.adcos.com.br/backend/kits/form/")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "kit")))

        # NOME KIT - Preencher nome do kit
        campo_kit = driver.find_element(By.NAME, "kit")
        campo_kit.click()
        campo_kit.clear()
        campo_kit.send_keys(nome_personalizado)

        # CODIGO DO KIT - Insere o código do kit
        campo_valor = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "codigo_barra"))
        )
        campo_valor.click()
        campo_valor.clear()
        campo_valor.send_keys(kit_codigo_formatado)

        #DATAS
        data_inicial = produtos[0]["data_inicial"]
        data_final = produtos[0]["data_final"]

        #DATA INICIAL
        pyperclip.copy(data_inicial)
        campo_data_inicial = driver.find_element(By.NAME, "data_inicial")
        campo_data_inicial.click()
        campo_data_inicial.send_keys(Keys.CONTROL, 'v')

        #HORA INICIAL
        campo_valor = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "hora_inicial"))
        )
        campo_valor.click()
        campo_valor.clear()
        campo_valor.send_keys(hora_inicial)

        #DATA FINAL
        pyperclip.copy(data_final)
        campo_data_final = driver.find_element(By.NAME, "data_final")
        campo_data_final.click()
        campo_data_final.send_keys(Keys.CONTROL, 'v')

        # HORA FINAL
        campo_valor = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "hora_final"))
        )
        campo_valor.click()
        campo_valor.clear()
        campo_valor.send_keys(hora_final)

        # BOTÃO PARA SALVAR
        driver.find_element(By.ID, "bt-salvar").click()
        time.sleep(2)

        # Espera a URL conter "/kits/form/" (garante que a página carregou completamente)
        WebDriverWait(driver, 20).until(EC.url_contains("/kits/form/"))
        url = driver.current_url

        # Captura o número após "/form/"
        match = re.search(r"/form/(\d+)", url)

        if match:
            numero = match.group(1)
            print("Número capturado:", numero)

            # Insere a URL na coluna I (coluna 9) na linha correspondente
            abaplanilha.Cells(linha_kit, 9).Value = numero
        else:
            print("Número do kit não encontrado na URL.")



        # ÁREAS PARA INSERIR OS PRODUTOS - Aqui vamos inserir todos os produtos que estão dentro do kit
        elemento = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@rel='box-produtos']/span[text()='Produtos']"))
        )

        elemento.click()

        # COLOCANDO TODOS OS PRODUTOS DO KIT
        for produto in produtos:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "produto_anexo")))

            #Digite o produto e espera o autocomplete
            campo_produto = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "produto_anexo")))
            campo_produto.clear()  
            campo_produto.send_keys(produto["produto"])
            time.sleep(1)  

            # Aguardar a sugestão aparecer e clicar nela
            try:
                sugestao = WebDriverWait(driver, 20).until(
                    EC.visibility_of_element_located((By.XPATH, "//ul[contains(@class, 'ui-autocomplete')]/li"))
                )

                # Rola a página até o elemento
                driver.execute_script("arguments[0].scrollIntoView();", sugestao)
                sugestao.click()

                # Esconde o autocomplete via JavaScript
                driver.execute_script("$('.ui-autocomplete').hide();")
                time.sleep(1)

                # Confirmar se o campo oculto foi preenchido corretamente
                hidden_input = driver.find_element(By.XPATH, "//input[@name='id_produto']")
                log(f"Produto {produto['produto']} adicionado com sucesso. ID: {hidden_input.get_attribute('value')}\n", log_area)

                time.sleep(1)

            # Passa para o próximo produto em caso de erro
            except Exception as e:
                log(f"ERROR: Erro ao adicionar o produto {produto['produto']}: {e}\n", log_area)
                continue 

            # QUANTIDADE - Preencher a quantidade
            campo_quantidade = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@name='quantidade']"))
            )
            campo_quantidade.clear()
            campo_quantidade.send_keys(quantidade)
            campo_quantidade.send_keys(Keys.RETURN)
            time.sleep(2)

            #Recarrega a pagina
            driver.refresh()

            # Reacessar os produtos para adicionar o próximo
            elemento = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@rel='box-produtos']/span[text()='Produtos']"))
            )
            elemento.click()

        # TABELA DE PRECOS - Aqui vamos inserir os preços do kit nas tabelas 

        #Print para verificar o preço total
        log(f"Total preço com desconto do kit R${total_preco_desconto:.2f}\n", log_area)
        log(f"Total preço padrão do kit R${total_preco_padrao:.2f}\n", log_area)

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
        time.sleep(1)

        #TABELA 00 PREÇO PADRÂO - Aqui vamos colocar o preço total sem desconto
        opcao = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='text' and text()='A00 - PRECO PADRAO']"))
        )
        opcao.click() 

        #FORMATANDO PREÇOS PARA INSERIR NO PW - Arredonda para 2 casa decimais
        total_preco_padrao = round(total_preco_padrao, 2) 
        total_preco_padrao = f"{total_preco_padrao:.2f}".replace('.', ',')

        total_preco_desconto = round(total_preco_desconto, 2)
        total_preco_desconto = f"{total_preco_desconto:.2f}".replace('.', ',')

        #VALOR A VISTA - Preço total padrão
        campo_preco_padrao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_vista"))
        )
        campo_preco_padrao.clear()
        campo_preco_padrao.send_keys(total_preco_padrao)

        #VALOR A PRAZO - Preço total padrão
        campo_preco_padrao1 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_prazo"))
        )

        campo_preco_padrao1.click()
        campo_preco_padrao1.clear()
        campo_preco_padrao1.send_keys(total_preco_padrao)
        campo_preco_padrao1.send_keys(Keys.RETURN)

        #Recarrega a pagina
        driver.refresh()

        #Entra na pagina de preços novamente
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
        time.sleep(1)

        #TABELA A19 - Cliente final, aqui vamos inserir com desconto
        opcao_2 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='text' and text()='A19 - CLIENTE FINAL']"))
        )
        opcao_2.click()  

        #VALOR A VISTA - Preço com desconto
        campo_preco_desconto = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_vista"))
        )
        campo_preco_desconto.clear()
        campo_preco_desconto.send_keys(total_preco_desconto)

        #VALOR A PRAZO - Preço com desconto
        campo_preco_desconto1 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "valor_prazo"))
        )

        campo_preco_desconto1.click()
        campo_preco_desconto1.clear()
        campo_preco_desconto1.send_keys(total_preco_desconto)
        campo_preco_desconto1.send_keys(Keys.RETURN)
        time.sleep(2)

    # FECHAR O NAVEGADOR - Encerrar a execução
    driver.quit()
    log("✅ Formulários preenchidos com sucesso!\n", log_area)
