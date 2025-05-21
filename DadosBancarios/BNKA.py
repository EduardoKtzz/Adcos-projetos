import win32com.client
import time
import sys

# Acessa a planilha aberta no Excel
excel = win32com.client.Dispatch("Excel.Application")
ws_dados = None

# Itera sobre todas as pastas de trabalho abertas
for ws_dados in excel.Workbooks:
    # Verifica se a planilha existe na pasta de trabalho atual
    try:
        ws_dados = ws_dados.Sheets("BNKA")  # Tenta acessar a aba pelo nome
        print(f"Planilha encontrada na pasta de trabalho '{ws_dados.Name}'")
        break  # Sai do loop se encontrar a planilha desejada
    except Exception as e:
        print(f"Planilha não encontrada na pasta de trabalho '{ws_dados.Name}'")

    # Se a planilha não foi encontrada, encerra o script
if not ws_dados:
    print("🚨 ERRO: A planilha não foi encontrada em nenhuma pasta de trabalho aberta.")
    sys.exit("Encerrando o robô...")

# Conecta ao SAP GUI
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Maximiza a janela
session.findById("wnd[0]").maximize()

# Acessa a transação SE16N
session.findById("wnd[0]/tbar[0]/okcd").text = "se16n"
session.findById("wnd[0]").sendVKey(0)

# Define a tabela BNKA
session.findById("wnd[0]/usr/ctxtGD-TAB").text = ""
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "BNKA"
session.findById("wnd[0]").sendVKey(0)

# Remove o limite de registros (campo MAX_LINES)
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]").sendVKey(0)

# Executa a busca (F8)
session.findById("wnd[0]").sendVKey(8)

# Exporta para Excel via menu de contexto
shell = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
shell.pressToolbarContextButton("&MB_VARIANT")
shell.setCurrentCell(2, "BANKA")
shell.contextMenu()
shell.selectContextMenuItem("&XXL")

# Nomeia e confirma a exportação
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "BNKA(EXPORT).XLSX"
session.findById("wnd[1]/tbar[0]/btn[11]").press()
