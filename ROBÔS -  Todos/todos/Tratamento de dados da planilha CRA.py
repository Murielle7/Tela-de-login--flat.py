import pandas as pd

# Caminho para o arquivo
caminho = 'S:/CAJ/Servidores/Murielle/relatorio-remessa(5)- dados tratados.xlsx'

# Lendo a planilha
planilha = pd.read_excel(caminho, sheet_name='Planilha1')

# Verificando e tratando as linhas da coluna 'Devedor' com múltiplas linhas
# Supondo que as células que possuem múltiplas linhas são da forma 'linha1\nlinha2'
def expand_rows(row):
    # Divide a coluna 'Devedor' em várias linhas
    devedores = row['Devedor'].split('\n')
    # Para cada devedor, retorna uma nova linha com os dados correspondentes de 'Número' e 'Nosso número'
    return pd.DataFrame({
        'Devedor': devedores,
        'Número': [row['Número']] * len(devedores),
        'Nosso número': [row['Nosso número']] * len(devedores),
        'Ocorrência': [row['Ocorrência']] * len(devedores)
    })

# Cria uma nova lista de DataFrame para cada linha expandida
new_rows = []

for _, row in planilha.iterrows():
    new_rows.append(expand_rows(row))

# Concatena todas as novas linhas em um DataFrame
result = pd.concat(new_rows, ignore_index=True)

# Exibindo o resultado
print(result)

# Se desejar, você pode salvar o resultado em um novo arquivo Excel
result.to_excel('S:/CAJ/Servidores/Murielle/Arquivo CRA tratado 2.xlsx', index=False)
