#Jogar dados na planilha 

# Aguarda o Excel abrir a planilha exportada pelo SAP
time.sleep(4)  # dá tempo de abrir

# CAÇA A PLANILHA - Verifica todas planilhas para achar certa de KITS
for abaplanilha in excel.Workbooks:
    try:
        abaplanilha = abaplanilha.Sheets("EXPORT")
        print(f"Planilha encontrada na pasta de trabalho '{abaplanilha.Name}'")
        break
    except Exception as e:
        print(f"Planilha não encontrada na pasta de trabalho '{abaplanilha.Name}'")

# Procurando a aba que o SAP geralmente exporta — ex: "Sheet1" ou "Planilha1"
for wb in excel.Workbooks:
    try:
        if wb.Sheets(1).Name.lower() in ["sheet1", "planilha1"]:
            planilha_exportada = wb
            print(f"✅ Planilha exportada encontrada: {wb.Name}")
            break
    except:
        continue

if not planilha_exportada:
    print("❌ Nenhuma planilha exportada foi localizada.")
    exit()

# Agora você pode manipular a planilha normalmente
aba = planilha_exportada.Sheets(1)
valor = aba.Cells(2, 1).Value  # exemplo: lê célula A2
print(f"📄 Valor na célula A2: {valor}")