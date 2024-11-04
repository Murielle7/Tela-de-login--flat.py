import os
import pandas as pd

# Caminho da pasta e do arquivo Excel
pasta_arquivos = "S:\\CAJ\\Servidores\\Murielle\\boletos\\Remessa 26-07"
caminho_excel = "S:/CAJ/Servidores/Murielle/X90J260724 Goiânia(1)- Usar para alterar status.xlsx"

# Ler a coluna CPFCNPJ do arquivo Excel
df = pd.read_excel(caminho_excel, usecols=['CPFCNPJ'])  # Assumindo que a coluna é 'CPFCNPJ'
cpfs_cnpjs = df['CPFCNPJ'].tolist()

# Listar os arquivos e extrair a segunda parte
for arquivo in os.listdir(pasta_arquivos):
    partes = arquivo.split("_")
    if len(partes) >= 2:
        segunda_parte = partes[1]
        if segunda_parte in cpfs_cnpjs:
            print(f"A segunda parte '{segunda_parte}' do arquivo '{arquivo}' foi encontrada na coluna CPFCNPJ do Excel.")
        else:
            print(f"A segunda parte '{segunda_parte}' do arquivo '{arquivo}' NÃO foi encontrada na coluna CPFCNPJ do Excel.")
    else:
        print(f"O arquivo {arquivo} não possui uma segunda parte.")
