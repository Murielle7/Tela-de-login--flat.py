import os

# Caminho da pasta que contém os arquivos
caminho_pasta_arquivos = r'S:/CAJ/Servidores/Murielle/Demanda GISELE -  Setembro  - SEM DUPLICATAS.xlsx'

# Listar todos os arquivos na pasta
arquivos_pasta = os.listdir(caminho_pasta_arquivos)

for nome_arquivo in arquivos_pasta:
    if os.path.isfile(os.path.join(caminho_pasta_arquivos, nome_arquivo)):
        nome_arquivo_atual = os.path.splitext(nome_arquivo)[0]
        extensao_arquivo = os.path.splitext(nome_arquivo)[1]
        novo_nome_arquivo = f"{nome_arquivo_atual}_01{extensao_arquivo}"
        os.rename(os.path.join(caminho_pasta_arquivos, nome_arquivo), os.path.join(caminho_pasta_arquivos, novo_nome_arquivo))

print("O sufixo '_01' foi adicionado a todos os arquivos da pasta.")
