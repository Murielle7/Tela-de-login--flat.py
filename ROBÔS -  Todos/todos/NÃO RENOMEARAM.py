import os
import shutil

# Define os caminhos das pastas
caminho_origem = r'S:\CAJ\Servidores\Murielle\Demanda CRA - Setembro'
caminho_destino = r'S:\CAJ\Servidores\Murielle\Demanda CRA - Setembro\INAPTOS'

# Cria a pasta de destino, se não existir
if not os.path.exists(caminho_destino):
    os.makedirs(caminho_destino)

# Lista todos os arquivos na pasta de origem
arquivos = os.listdir(caminho_origem)

# Verifique se está lendo os arquivos corretamente
print(f"Arquivos encontrados na pasta de origem: {arquivos}")

# Percorre todos os arquivos na pasta de origem
for arquivo in arquivos:
    caminho_arquivo = os.path.join(caminho_origem, arquivo)
    
    # Verifica se é realmente um arquivo
    if os.path.isfile(caminho_arquivo):
        # Divide o nome do arquivo em partes usando pontos como separador para identificar o nome base
        partes_nome = arquivo.split('.')

        # Verifica se o nome base do arquivo tem apenas uma parte (sem sublinhados)
        if len(partes_nome) > 1:
            nome_base = partes_nome[0]
            if '_' not in nome_base:
                # Move o arquivo para a pasta de destino
                try:
                    caminho_destino_arquivo = os.path.join(caminho_destino, arquivo)
                    shutil.move(caminho_arquivo, caminho_destino_arquivo)
                    print(f'Movido: {arquivo}')
                except Exception as e:
                    print(f'Erro ao mover {arquivo}: {e}')
        else:
            print(f"Ignorado (não tem nenhuma extensão): {arquivo}")

print("Processo concluído.")
