import os
import shutil

caminho_origem = r"C:\Users\mferreirasantos\Downloads"
caminho_destino = r"S:\CAJ\Cobrança - Judicial e Administrativa\Custas Judiciais - Finais\Remessas - CRA\Boleto_Remessa"

# Verifica se o diretório de origem existe
if os.path.isdir(caminho_origem):
    # Lista todos os arquivos na pasta de origem
    arquivos = os.listdir(caminho_origem)

    # Itera sobre os arquivos e os move para a pasta de destino, se forem arquivos PDF
    for arquivo in arquivos:
        caminho_arquivo_origem = os.path.join(caminho_origem, arquivo)
        
        # Verifica se é um arquivo PDF
        if os.path.isfile(caminho_arquivo_origem) and arquivo.lower().endswith('.pdf'):
            caminho_arquivo_destino = os.path.join(caminho_destino, arquivo)
            shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)
            print(f"Arquivo {arquivo} movido para {caminho_destino}")
else:
    print("Diretório de origem não encontrado.")
