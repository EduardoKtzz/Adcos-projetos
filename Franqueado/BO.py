#Importações para acessar o excel e colocar tempo nas execuções
import win32com.client
import time

#Conecta em algum excel que estiver aberto e o nome da aba tem que ser o mesmo da linha 8
excel = win32com.client.GetActiveObject("Excel.Application")
wb = excel.ActiveWorkbook
ws = wb.Sheets("BO") 

#Dados para conectar no SAP
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

#Maximiza janela do SAP
session.findById("wnd[0]").maximize

#Aqui escolhemos o número da linha que o codigo vai começar a executar na planilha
linha = 3

#Aqui vamos colocar a estrutura de repitição para pegar todas linhas preenchidas da planilha
while True:
    bp = str(int(ws.Cells(linha, 1).Value)) #BP do cliente

    print(f"Inserindo email para BP {bp}...")

    #Acessando a transação BP e colocando o BP do cliente
    session.findById("wnd[0]/tbar[0]/okcd").text = "bp"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/ctxtBUS_JOEL_MAIN-OPEN_NUMBER").text = str(bp)
    session.findById("wnd[1]/usr/ctxtBUS_JOEL_MAIN-OPEN_NUMBER").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                 "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                 "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                 "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                 "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04").select()

    session.findById("wnd[0]/tbar[1]/btn[6]").press()

    path = ("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
            "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
            "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
            "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
            "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04/"
            "ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/"
            "ssubGENSUB:SAPLBUSS:7028/subA02P03:SAPLBUD0:1050/ctxtBUT000-AUGRP")

    session.findById(path).text = ""
    session.findById(path).setFocus()
    session.findById(path).caretPosition = 0

    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[6]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                    "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                    "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                    "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/"
                    "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01").select()

    #Marcando na coluna C o resultado do codigo
    ws.Cells(linha, 3).Value = "Sucesso"
    print(f"Cliente {bp} atualizado com sucesso.")

    time.sleep(1)

    #Vai para a próxima linha
    linha += 1