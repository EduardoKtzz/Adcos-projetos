#IMPORTAÇÔES PARA O CODIGO FUNCIONAR
import win32com.client
from datetime import date, datetime
import re

#CONFIGURAÇÃO PARA O SAP ABRIR E FUNCIONAR 
SapGuiAuto = win32com.client.GetObject("SAPGUI")
Appl = SapGuiAuto.GetScriptingEngine
Connection = Appl.Children(0)
session = Connection.Children(0)

#PEGAR A DATA ATUAL DA EXECUÇÃO DO SCRIPT
data_original = date.today()
data_formatada = data_original.strftime("%d.%m.%Y")

#CAMPO PARA PEGAR O CPF DO ESTUDANTE
cpf_estudante = input("Digite o CPF do estudante: ")

#FORMATAR O CPF - TIRAR PONTOS E ESPAÇOS
cpf_estudante = re.sub(r"\D", "", cpf_estudante)

#ABRE ROTINA DE BP
session.findById("wnd[0]").maximize()	
session.findById("wnd[0]/tbar[0]/okcd").Text = "bp"
session.findById("wnd[0]").sendVKey(0)

#COLETA AS INFORMAÇÕES DO COLABORADOR
session.findById("wnd[0]/tbar[1]/btn[6]").press()	
nome_estudante = session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/txtBUS_JOEL_MAIN-CHANGE_DESCRIPTION").text		
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/txtBUS_JOEL_MAIN-CHANGE_DESCRIPTION").caretPosition = 32
session.findById("wnd[0]").sendVKey(2)
print(nome_estudante)

#PESQUISAR O CPF E SALVAR O NUMERO DE BP 	
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2120/txtBUS_JOEL_SEARCH-EXTERNAL_NUMBER").Text = cpf_estudante
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2120/txtBUS_JOEL_SEARCH-EXTERNAL_NUMBER").caretPosition = 11
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_BUTTON_AREA:SAPLBUS_LOCATOR:3240/btnBUS_LOCA_SRCH01-GO").press()
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell").currentCellColumn = "DESCRIPTION"
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1060/ssubSCREEN_1060_RESULT_AREA:SAPLBUPA_DIALOG_JOEL:1080/cntlSCREEN_1080_CONTAINER/shellcont/shell").doubleClickCurrentCell()
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

status_analise = session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE/ctxtSTRUC-FIELD3-VALUE-LEFT[3,0]").text
print('Cadastro está:', status_analise)

# Verificar o valor do status
if status_analise == "V":
    print("O status é 'V' (vencido). Continuando o cadastro como estudante.")
    
    # Realizar ações para cadastrar como estudante
    #PRIMEIRA TELA DO MASS
    #AQUI VAMOS COLOCAR O 119(CLIENTE FINAL), MARCAR O X DE PESSOA FISICA E COLOCAR O A DE APROVADO
    #E POR ULTIMO VAMOS CLICAR NO ITEM PARA DEIXAR TODOS IGUAIS
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD2-VALUE-LEFT[2,0]").text = "119"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD3-VALUE-LEFT[3,0]").text = "A"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD4-VALUE-LEFT[4,0]").text = "X"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press()

    #AQUI ELE VAI PARA A TELA DO LADO DO MASS
    #AQUI VAMOS MARCAR ELE COM SEGMENTO E SUBSEGUIMENTO DE 05 e 051 E DEPOIS VAMOS CLICAR NO ITEM PARA DEIXAR TUDO IGUAL
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2").select()
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[5,0]").text = "11"
    session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").text = "119"
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
    session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/txtDRAT-DKTXT").text = "DOCUMENTO PROFISSIONAL " + data_formatada

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
        
    print("Cadastro concluído com sucesso!")

elif status_analise.strip() == "P":

    # Perguntar ao usuário como deseja proceder
    novo_status = input("O status está vazio. Deseja inserir um novo valor? (S/N): ").strip().upper()
    if novo_status == "S":
        valor_inserir = input("Digite o novo valor para o status: ").strip()
        
        # Atualizar o campo no SAP
        #PRIMEIRA TELA DO MASS
        #AQUI VAMOS COLOCAR O 119(CLIENTE FINAL), MARCAR O X DE PESSOA FISICA E COLOCAR O A DE APROVADO
        #E POR ULTIMO VAMOS CLICAR NO ITEM PARA DEIXAR TODOS IGUAIS
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD2-VALUE-LEFT[2,0]").text = valor_inserir
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD3-VALUE-LEFT[3,0]").text = "A"
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD4-VALUE-LEFT[4,0]").text = "X"
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press()

        #AQUI ELE VAI PARA A TELA DO LADO DO MASS
        #AQUI VAMOS MARCAR ELE COM SEGMENTO E SUBSEGUIMENTO DE 05 e 051 E DEPOIS VAMOS CLICAR NO ITEM PARA DEIXAR TUDO IGUAL
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2").select()
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[5,0]").text = "11"
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").text = valor_inserir
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").setFocus()
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD6-VALUE-LEFT[6,0]").caretPosition = 3
        session.findById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB2/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press()
        print(f"Novo valor '{valor_inserir}' inserido com sucesso!")

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
        session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/txtDRAT-DKTXT").text = "DOCUMENTO PROFISSIONAL " + data_formatada

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
            
        print("Cadastro concluído com sucesso!")

    else:
        print("Operação cancelada pelo usuário.")