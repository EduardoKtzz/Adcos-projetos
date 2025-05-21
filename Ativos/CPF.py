import re
import win32com.client
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

# NAO ACHAR PLANILHA - Se a planilha não foi encontrada, encerra o script
if not abaplanilha:
    print("🚨 ERRO: A planilha 'Cadastro de Kits' não foi encontrada em nenhuma pasta de trabalho aberta.")
    sys.exit("Encerrando o robô...")

# FORMATA A COLUNA E COMO TEXTO
abaplanilha.Columns(5).NumberFormat = "@"  # "@" = Texto

# TOTAL DE LINHA - Descobre o número total de linhas preenchidas na planilha
ultima_linha = abaplanilha.Cells(abaplanilha.Rows.Count, 4).End(-4162).Row

# LOOP PARA FORMATAR OS CPFs DA COLUNA D E JOGAR NA COLUNA E
for linha in range(2, ultima_linha + 1):  # Começa da linha 2
    valor = abaplanilha.Cells(linha, 4).Value
    if valor:
        cpf_original = str(valor).strip()
        cpf_limpo = re.sub(r'[^\d]', '', cpf_original)  # Remove tudo que não for número
        cpf_limpo = cpf_limpo.zfill(11)  # Garante 11 dígitos com zeros à esquerda
        abaplanilha.Cells(linha, 5).Value = cpf_limpo  # Escreve na coluna E

print("✅ CPFs formatados com sucesso e inseridos na coluna E.")
