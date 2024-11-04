import os
from collections import defaultdict

# Defina o caminho da pasta
caminho_pasta = r'S:\CAJ\Servidores\Murielle\Demanda CRA - Setembro'

# Cria um dicionário para armazenar os arquivos de cada "primeira parte"
arquivos_por_parte = defaultdict(list)

# Percorre todos os arquivos na pasta
for arquivo in os.listdir(caminho_pasta):
    # Caminho completo do arquivo
    caminho_arquivo = os.path.join(caminho_pasta, arquivo)
    
    # Ignora se não for arquivo
    if not os.path.isfile(caminho_arquivo):
        continue
    
    # Divide o nome do arquivo na primeira ocorrência de "_" ou "." para capturar a "primeira parte"
    primeira_parte = arquivo.split('_')[0].split('.')[0]
    
    # Adiciona o arquivo à lista correspondente à sua primeira parte
    arquivos_por_parte[primeira_parte].append(arquivo)

# Itera sobre cada grupo de arquivos com a mesma "primeira parte"
for arquivos in arquivos_por_parte.values():
    # Se houver mais de um arquivo, mantém apenas o primeiro e apaga os outros
    if len(arquivos) > 1:
        # Mantém o primeiro arquivo e remove os demais
        for arquivo_para_remover in arquivos[1:]:
            # Caminho completo do arquivo a ser removido
            caminho_arquivo_para_remover = os.path.join(caminho_pasta, arquivo_para_remover)
            try:
                os.remove(caminho_arquivo_para_remover)
                print(f"Removido: {arquivo_para_remover}")
            except Exception as e:
                print(f"Erro ao remover {arquivo_para_remover}: {e}")

print("Processo concluído. Apenas um arquivo por grupo de repetição foi mantido.")
