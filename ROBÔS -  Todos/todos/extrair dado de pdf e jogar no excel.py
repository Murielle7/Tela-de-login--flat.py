import fitz  # PyMuPDF
import re
import pandas as pd

# Função para encontrar as datas no texto
def encontrar_datas(texto):
    # Expressão regular para datas no formato dd/mm/aaaa
    padrao_data = r'\b\d{2}/\d{2}/\d{4}\b'
    datas = re.findall(padrao_data, texto)
    return datas

# Caminho do arquivo PDF
caminho_pdf = 'C:/Users/mferreirasantos/Downloads/GuiaFinal_FinalZeroProcesso_null_05_08_2024 13_15_081722874508950.pdf'

# Abrir e ler o PDF
documento_pdf = fitz.open(caminho_pdf)
texto_pdf = ""
for pagina in documento_pdf:
    texto_pdf += pagina.get_text()

# Encontrar as datas no texto
datas_encontradas = encontrar_datas(texto_pdf)

# Verificar se há pelo menos duas datas encontradas
if len(datas_encontradas) >= 2:
    data_vencimento = datas_encontradas[1]  # Pegando a segunda data encontrada
else:
    data_vencimento = "Data não encontrada"

# Fechar o documento PDF
documento_pdf.close()

# Caminho do arquivo Excel
caminho_excel = 'S:/CAJ/Servidores/Murielle/Teste - dem. Pedro.xlsx'

# Ler a planilha 'Sheet1'
xls = pd.ExcelFile(caminho_excel, engine='openpyxl')
df = pd.read_excel(xls, sheet_name='Sheet1')

# Verificar se a coluna 'Data Vencimento Original' existe
if 'Data Vencimento Original' in df.columns:
    # Atualizar apenas a célula na 4ª linha (index 3) sem alterar os outros dados
    df.at[3, 'Data Vencimento Original'] = data_vencimento
else:
    # Se a coluna 'Data Vencimento Original' não existe, crie a coluna e adicione a data na 4ª linha
    df['Data Vencimento Original'] = pd.Series([None] * len(df))  # Adiciona a nova coluna
    df.at[3, 'Data Vencimento Original'] = data_vencimento

# Salvar o arquivo Excel na planilha 'Sheet1'
with pd.ExcelWriter(caminho_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)

print("Arquivo salvo com sucesso.")
