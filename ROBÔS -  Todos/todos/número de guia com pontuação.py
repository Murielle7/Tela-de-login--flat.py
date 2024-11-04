import pandas as pd

# Carrega a planilha do Excel
file_path = "C:/Users/mferreirasantos/Documents/VERDE - FALTARAM parte 2 --- P2.xlsx"
sheet_name = "Planilha1"

# Lê o arquivo Excel especificando a planilha
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Função para formatar a string conforme solicitado
def format_string(value):
    value_str = str(value)
    if len(value_str) >= 3:
        formatted_str = value_str[:-3] + '-' + value_str[-3:-2] + '/' + value_str[-2:]
        return formatted_str
    return value_str  # Retorna sem alteração se a string for curta

# Aplica a função na coluna 'd'
df['Número Guia'] = df['Número Guia'].apply(format_string)

# Salva as alterações de volta no arquivo Excel
df.to_excel(file_path, sheet_name=sheet_name, index=False)

print("A formatação foi concluída com sucesso.")
