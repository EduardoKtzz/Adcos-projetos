import win32com.client
import win32com.client as win32
import pyperclip
import time
import sys


# ABRE EXCEL - Acessa a planilha aberta no Excel
excel = win32com.client.Dispatch("Excel.Application")
abaplanilha = None

# CAÇA A PLANILHA - Verifica todas planilhas para achar certa de KITS
for abaplanilha in excel.Workbooks:
    try:
        abaplanilha = abaplanilha.Sheets("AJUSTE_RH")
        print(f"Planilha encontrada na pasta de trabalho '{abaplanilha.Name}'")
        break
    except Exception as e:
        print(f"Planilha não encontrada na pasta de trabalho '{abaplanilha.Name}'")

# TOTAL DE LINHA - Descobre o número total de linhas preenchidas na planilha
ultima_linha = abaplanilha.Cells(abaplanilha.Rows.Count, 5).End(-4162).Row

# COPIA OS VALORES PARA A LISTA
valores = []
for linha in range(2, ultima_linha + 1):  # Começa da linha 2
    valor = abaplanilha.Cells(linha, 5).Value
    if valor:
        valores.append(str(valor).strip())

# COPIA PRA ÁREA DE TRANSFERÊNCIA — um por linha (usando \r\n que o SAP ama)
pyperclip.copy("\r\n".join(valores))

print(f"📋 {len(valores)} CPF(s) copiado(s) para a área de transferência.")

# Conecta ao SAP GUI
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# (Opcional) Maximizar janela
session.findById("wnd[0]").maximize()

# Navega até a transação SQVI
session.findById("wnd[0]/tbar[0]/okcd").text = "sqvi"
session.findById("wnd[0]").sendVKey(0)

#Verificar se a opção está marcada
tabela = session.findById("wnd[0]/usr/tblSAPMS38RTV3050")
linha = tabela.getAbsoluteRow(0)
linha.selected = True

# Interações dentro do SQVI
session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,0]").setFocus()
session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,0]").caretPosition = 0
session.findById("wnd[0]/usr/btnP1").press()

# Seleciona valores de entrada
session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00001-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()

# Seleciona célula da tabela
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")

# Confirma exportação
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()


