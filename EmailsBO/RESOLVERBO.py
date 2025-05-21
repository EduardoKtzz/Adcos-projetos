#Importações para acessar o excel e colocar tempo nas execuções
import win32com.client
import time

#Conecta em algum excel que estiver aberto e o nome da aba tem que ser o mesmo da linha 8
excel = win32com.client.GetActiveObject("Excel.Application")
wb = excel.ActiveWorkbook
ws = wb.Sheets("Planilha1") 

#Dados para conectar no SAP
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

#Maximiza janela do SAP
session.findById("wnd[0]").maximize()

#Aqui escolhemos o número da linha que o codigo vai começar a executar na planilha
linha = 2

#Aqui vamos colocar a estrutura de repitição para pegar todas linhas preenchidas da planilha
while True:
    bp = str(int(ws.Cells(linha, 1).Value)) #BP do cliente
    email = ws.Cells(linha, 2).Value #Email novo do cliente

    #Caso a linha não tenha BP ou email, o codigo para de excutar
    if not bp or not email:
        break  

    print(f"Inserindo email para BP {bp}...")

    #Acessando a transação BP e colocando o BP do cliente
    session.findById("wnd[0]/tbar[0]/okcd").text = "bp"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/ctxtBUS_JOEL_MAIN-OPEN_NUMBER").text = str(bp)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #Ativando o modo edição
    session.findById("wnd[0]/tbar[1]/btn[6]").press()
    time.sleep(1.5)

    #Insere o email (1º campo)
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtSZA7_D0400-SMTP_ADDR").text = email

    #Insere o email (2º campo)
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA06P01:SAPLBUA0:0700/subADDR_ICOMM:SAPLSZA11:0100/txtSZA11_0100-SMTP_ADDR").text = email

    # Salva clicando no botão de desativar edição
    session.findById("wnd[0]/tbar[1]/btn[6]").press()

    # Fecha popup de informação clicando em OK e caso não tenha ele continua o caso
    try:
        while session.findById("wnd[1]").Text == "Informação":
            session.findById("wnd[1]/tbar[0]/btn[0]").press()  
            time.sleep(0.5)
    except:
        pass

    #Aqui ele confirma que quer salvar as alterações
    try:
        if session.findById("wnd[1]/usr/btnSPOP-OPTION1").Text in ["Sim", "Yes"]:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    except:
        pass

    #Marcando na coluna C o resultado do codigo
    ws.Cells(linha, 3).Value = "Sucesso"
    print(f"Cliente {bp} atualizado com sucesso.")

    time.sleep(1)

    #Vai para a próxima linha
    linha += 1
