import os

pasta = r"S:\CAJ\Servidores\Murielle\teste remessa 26-07"

# Função para renomear arquivos
def renomear_arquivos(pasta):
    # Percorre todos os arquivos na pasta
    for nome_arquivo in os.listdir(pasta):
        caminho_atual = os.path.join(pasta, nome_arquivo)
        
        # Verifica se é um arquivo
        if os.path.isfile(caminho_atual):
            # Separa o nome do arquivo e a extensão
            nome_base, extensao = os.path.splitext(nome_arquivo)
            
            # Separa o nome do arquivo por underscores
            partes = nome_base.split('_')
            
            # Verifica se há pelo menos 5 partes separadas por underscores
            if len(partes)==2:
                # Remove a 5ª sequência de números (índice 4) das partes
                partes.pop(1)
                
                # Junta as partes novamente com underscores
                novo_nome_base = '_'.join(partes)
                novo_nome_arquivo = novo_nome_base + extensao
                
                # Novo caminho do arquivo com o novo nome
                novo_caminho = os.path.join(pasta, novo_nome_arquivo)
                
                # Renomeia o arquivo
                os.rename(caminho_atual, novo_caminho)
                print(f"Arquivo renomeado: {nome_arquivo} -> {novo_nome_arquivo}")

# Chama a função para renomear os arquivos na pasta especificada
renomear_arquivos(pasta)

