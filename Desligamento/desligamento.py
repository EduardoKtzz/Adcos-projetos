import tkinter as tk
from tkinter import Button, Entry, Label, messagebox
import win32com.client
import re
from datetime import date
import ctypes  # Melhorar a qualidade das imagens 
import customtkinter


def sap(cpf_input, resultado_label):
	#CONFIGURAÇÃO PARA O SAP ABRIR E FUNCIONAR 
	SapGuiAuto = win32com.client.GetObject("SAPGUI")
	Appl = SapGuiAuto.GetScriptingEngine
	Connection = Appl.Children(0)
	session = Connection.Children(0)

	#PEGAR A DATA ATUAL DA EXECUÇÃO DO SCRIPT
	data_original = date.today()
	data_formatada = data_original.strftime("%d.%m.%Y")

	#FORMATAR O CPF - TIRAR PONTOS E ESPAÇOS
	cpf_colaborador = re.sub(r"\D", "", cpf_input)

	#ABRE ROTINA DE BP
	session.findById("wnd[0]").maximize()	
	session.findById("wnd[0]/tbar[0]/okcd").Text = "bp"
	session.findById("wnd[0]").sendVKey(0)

	#COLETA AS INFORMAÇÕES DO COLABORADOR
	session.findById("wnd[0]/tbar[1]/btn[6]").press()
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/txtBUS_JOEL_MAIN-CHANGE_DESCRIPTION").caretPosition = 32
	session.findById("wnd[0]").sendVKey(2)

	#PESQUISAR O CPF E SALVAR O NUMERO DE BP 	
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2120/txtBUS_JOEL_SEARCH-EXTERNAL_NUMBER").Text = cpf_colaborador
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2120/txtBUS_JOEL_SEARCH-EXTERNAL_NUMBER").caretPosition = 11
	session.findById("wnd[0]").sendVKey(0)
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_BUTTON_AREA:SAPLBUS_LOCATOR:3240/btnBUS_LOCA_SRCH01-GO").press()
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell").currentCellColumn = "DESCRIPTION"
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell").selectedRows = "0"
	session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell").doubleClickCurrentCell()
	nome_colaborador = session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/txtBUS_JOEL_MAIN-CHANGE_DESCRIPTION").text		
	BPNumber = session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/ctxtBUS_JOEL_MAIN-CHANGE_NUMBER").Text

	#MASS
	#DIGITAR A TRANSAÇÃO MASS
	session.findById("wnd[0]/tbar[0]/btn[3]").press()
	session.findById("wnd[0]/tbar[0]/okcd").text = "mass"
	session.findById("wnd[0]").sendVKey(0)

	#DIGITAR AS TRANSAÇÔES KNA1 e VAL_VETEX
	session.findById("wnd[0]/usr/ctxtMASSSCREEN-OBJECT").text = "KNA1"
	session.findById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").text = "VAL_VETEX"
	session.findById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").setFocus()
	session.findById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").caretPosition = 9
	session.findById("wnd[0]/tbar[1]/btn[8]").press()

	#COLOCAR O BP PARA EDIÇÂO NO MASS
	session.findById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").text = BPNumber
	session.findById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").caretPosition = 6
	session.findById("wnd[0]/tbar[1]/btn[8]").press()

	#PRIMEIRA TELA DO MASS
	#AQUI VAMOS VERIFICAR QUAL O SEGMENTO E O SUBSEGMENTO DO COLABORADOR, SE FOR 09 e 091 O CODIGO VAI PROSSEGUIR, SE FOR OUTRO O CODIGO PARA
	session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2").select()		
	segmento = session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD5-VALUE-LEFT[5,0]").text.strip()	
	subsegmento = session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD6-VALUE-LEFT[6,0]").text.strip()		
	franqueado = session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD7-VALUE-LEFT[7,0]").text.strip()
	session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1").select()

	#FAZ A VERIFICAÇÃO DO SEGMENTO DO COLABORADOR, SE FOR 091 ELE COLOCA COMO CLIENTE FINAL, CASO SEJA PROFISSIONAL ELE INTERROMPE O CODIGO
    #VERIFICA SE ELE É FRANQUEADO 
	if segmento == "09" and subsegmento == "091" and franqueado == "":
			mensagem_erro = ""
			#COLOCA COMO CLIENTE FINAL NA PRIMEIRA ABA DO MASS
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1").select()		
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD2-VALUE-LEFT[2,0]").text = "051"
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD4-VALUE-LEFT[4,0]").text = "X"
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD4-VALUE-LEFT[4,0]").setFocus()
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD4-VALUE-LEFT[4,0]").caretPosition = 1
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press()
			
			#COLOCA COMO CLIENTE FINAL NA SEGUNDA ABA DO MASS
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2").select()
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[5,0]").text = "05"
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").text = "051"
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").setFocus()
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").caretPosition = 3
			session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press()

			#PARA VOLTAR PARA A TELA INICIAL DO SAP
			session.findById("wnd[0]/tbar[0]/btn[11]").press()
			session.findById("wnd[0]/tbar[0]/btn[3]").press()
			session.findById("wnd[0]/tbar[0]/btn[3]").press()
			session.findById("wnd[0]/tbar[0]/btn[3]").press()
			session.findById("wnd[0]/tbar[0]/btn[3]").press()

			# SALVAR ITEM NO CV01N
			session.findById("wnd[0]/tbar[0]/okcd").text = "cv01n"
			session.findById("wnd[0]").sendVKey(0)

			#CRIAR DOCUMENTO COMO ARQUIVO 2
			session.findById("wnd[0]/usr/ctxtDRAW-DOKNR").Text = BPNumber
			session.findById("wnd[0]/usr/ctxtDRAW-DOKAR").Text = "ZXD"
			session.findById("wnd[0]/usr/ctxtDRAW-DOKTL").text = "002"
			session.findById("wnd[0]/usr/ctxtDRAW-DOKTL").SetFocus()
			session.findById("wnd[0]/usr/ctxtDRAW-DOKTL").caretPosition = 3
			session.findById("wnd[0]").sendVKey(0)

			#DEFINIR O TITULO COMO DESLIGAMENTO + DATA ATUAL 
			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/txtDRAT-DKTXT").text = "DESLIGAMENTO " + data_formatada

			# ABA DE LIGAÇÕES DE OBJETOS
			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSLINKS").select()
			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSLINKS/ssubSCR_MAIN:SAPLCV130:0404/tabsSELECT_OBJLINKSTRIP_400/tabpOBJTB01/ssubSUBSCRN_OBJLINK:SAPLCV130:1216/tblSAPLCV130TAB_X/ctxtKNA1-KUNNR[0,0]").text = BPNumber
			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSLINKS/ssubSCR_MAIN:SAPLCV130:0404/tabsSELECT_OBJLINKSTRIP_400/tabpOBJTB01/ssubSUBSCRN_OBJLINK:SAPLCV130:1216/tblSAPLCV130TAB_X/ctxtKNA1-KUNNR[0,0]").caretPosition = 10
			session.findById("wnd[0]").sendVKey(0)

			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN").Select()

			#PROCURA O ARQUIVO PARA SALVAR 
			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_BROWSER").press()
			session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press()
			session.findById("wnd[0]/tbar[0]/btn[11]").press()
    
	#AQUI ELE DIZ SE O SEGMENTO ESTIVER DIFERENTE OU SE ELE FOR FRANQUEADO
	else:
		mensagem_erro = []
		if segmento != "09" or subsegmento != "091":
			mensagem_erro.append(f"Segmento={segmento} ou Subsegmento={subsegmento} estão incorretos.")
		if franqueado != "":
			mensagem_erro.append(f"A aba Franqueado não está vazia: Franqueado='{franqueado}'.")

	#MENSAGEM INFORMATIVA SOBRE O PROCESSO
	if mensagem_erro:
		resultado = f"Nome: {nome_colaborador}\nBP: {BPNumber}\nErros:\n" + "\n".join(mensagem_erro)
	else:
		resultado = f"Nome: {nome_colaborador}\nBP: {BPNumber}\nTudo está correto, processo finalizado"

	resultado_label.configure(text=resultado)

#FUNÇÃO PARA VALIDAR O INPUT DE CPF
def validar_cpf(P):
    
    if P.isdigit() and len(P) <= 15:  # Apenas números e no máximo 15 dígitos
        return True
    elif P == "":  # Permitir campo vazio
        return True
    return False

#FUNÇÃO PARA INTERFACE 
def criar_interface():
	#CONFIGURAÇÂO INICIAL DA JANELA
	customtkinter.set_appearance_mode("System")
	customtkinter.set_default_color_theme("blue")
	menu = customtkinter.CTk()  # inicializa uma interface
	menu.title('Desligamento de Colaborador Adcos')  # titulo
	menu.geometry('520x270') # tamanho
	ctypes.windll.shcore.SetProcessDpiAwareness(1)

	validacao_cpf = menu.register(validar_cpf)

	#CONFIGURAÇÂO PARA CENTRALIZAR TEXTO
	menu.columnconfigure(0, weight=1)  # Centraliza horizontalmente
	menu.rowconfigure(0, weight=1)  # Linha superior ocupa espaço extra
	menu.rowconfigure(2, weight=3)  # Linha inferior ocupa espaço extra
	menu.configure(bg="#f0f0f0")

	# TEXTO DA LOGO
	texto_logo = customtkinter.CTkLabel(menu, text='Central de Cadastro Adcos - CCAD', font=('Helvetica', 16, "bold"))
	texto_logo.pack(pady=10)

	# TEXTO INFORMANDO PARA COLOCAR O CPF
	texto_explicativo = customtkinter.CTkLabel(menu, text='Digire o CPF do colaborador para realizar o desligamento:', font=('Helvetica', 14))
	texto_explicativo.pack(pady=5)

	# INPUT DO CPF
	cpf_input = customtkinter.CTkEntry(menu, width=300, validatecommand=(validacao_cpf, '%P'))
	cpf_input.pack(pady=5)
	
	# BOTÂO PARA EXECUTAR O SCRIPT
	botao_executar = customtkinter.CTkButton(menu, text='Começar desligamento', command=lambda:sap(cpf_input.get(), resultado_label))
	botao_executar.pack(pady=10)
	
	# LABEL PARA EXIBIR O RESULTADO FINAL
	resultado_label = customtkinter.CTkLabel(menu, text='', anchor="w", justify="left", font=("Arial", 12), padx=10, pady=5)
	resultado_label.pack(pady=10)

	resultado_label.configure(text_color="white", corner_radius=10)

	menu.mainloop() # finaliza a janela assim que clicado no X

if __name__ == "__main__":
    criar_interface()