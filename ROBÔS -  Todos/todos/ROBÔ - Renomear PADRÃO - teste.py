import os
import pandas as pd
import re

# Caminho do arquivo Excel e da pasta
arquivo_excel = r"S:/CAJ/Servidores/Murielle/remessa 20 --- 30-07.xlsx"
pasta = r"S:\CAJ\Servidores\Murielle\Remessa 20 - 30-07"

# Função para ler os dados das colunas CPF/CNPJ e Sequencial do arquivo Excel
def ler_colunas_excel(caminho_arquivo, nome_planilha, colunas):
    df = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, dtype=str, skiprows=3)  # Ignora as 3 primeiras linhas
    df.columns = [col.strip() for col in df.columns]  # Remover espaços extras dos nomes das colunas
    print("Colunas disponíveis:", df.columns)  # Imprimir nomes das colunas para depuração
    df = df[colunas].dropna(subset=colunas)
    
    # Remover caracteres não numéricos das colunas
    for coluna in colunas:
        df[coluna] = df[coluna].apply(lambda x: re.sub(r'\D', '', x))
    
    # Formatar o campo CPF/CNPJ para manter 11 dígitos para CPF e 14 dígitos para CNPJ
    def formatar_cpf_cnpj(valor):
        num_digits = len(valor)
        if num_digits <= 11:
            return valor.zfill(11)  # CPF com 11 dígitos
        elif num_digits <= 14:
            return valor.zfill(14)  # CNPJ com 14 dígitos
        else:
            return valor

    df['CPF/CNPJ'] = df['CPF/CNPJ'].apply(formatar_cpf_cnpj)
    
    # Verificar se todos os valores foram formatados corretamente
    df['CPF/CNPJ'] = df['CPF/CNPJ'].apply(lambda x: x if len(x) in {11, 14} else None)
    if df['CPF/CNPJ'].isnull().any():
        print("Erro de formatação nos seguintes valores:", df[df['CPF/CNPJ'].isnull()])

    return df

# Função para renomear os arquivos adicionando os sufixos do CPF/CNPJ e Sequencial
def renomear_arquivos_com_sufixo(pasta, sufixos):
    arquivos = os.listdir(pasta)
    for i, nome_arquivo in enumerate(arquivos):
        caminho_atual = os.path.join(pasta, nome_arquivo)
        
        # Verifica se é um arquivo e se há um sufixo disponível
        if os.path.isfile(caminho_atual) and i < len(sufixos):
            # Separa o nome do arquivo e a extensão
            nome_base, extensao = os.path.splitext(nome_arquivo)

            # Remove prefixo 'Boleto_TJGO_Guia_'
            if nome_base.startswith("Boleto_TJGO_Guia_"):
                nome_base = nome_base.replace("Boleto_TJGO_Guia_", "")
            
            # Divide o nome base em partes
            partes = nome_base.split('_')
            
            # Remove a 5ª parte se houver pelo menos 5 partes
            if len(partes) >= 5: #Ir subtraindo, e utilizar 5,4 e 3, para excluir partes corretas
                partes.pop(4) # Colocar o zero no lugar do 4, para retirar as partes não numéricas do nome (as primeiras partes do nome do arquivo)

            # Formata o sequencial para ter sempre 2 dígitos
            sequencial_formatado = sufixos[i]['Sequencial'].zfill(2)
            
            # Reconstrói o nome base
            novo_nome_base = '_'.join(partes)
            
            # Adiciona os sufixos ao nome do arquivo: CPF_CNPJ_Sequencial
            novo_nome_base = f"{novo_nome_base}_{sufixos[i]['CPF/CNPJ']}_{sequencial_formatado}"
            novo_nome_arquivo = f"{novo_nome_base}{extensao}"
            
            # Novo caminho do arquivo com o novo nome
            novo_caminho = os.path.join(pasta, novo_nome_arquivo)
            
            # Renomeia o arquivo
            os.rename(caminho_atual, novo_caminho)
            print(f"Arquivo renomeado: {nome_arquivo} -> {novo_nome_arquivo}")

# Lê as colunas CPF/CNPJ e Sequencial do arquivo Excel
colunas = ['CPF/CNPJ', 'Sequencial']
dados = ler_colunas_excel(arquivo_excel, 'Conferência - GO', colunas)

# Transforma os dados em uma lista de dicionários
sufixos = dados.to_dict('records')

# Chama a função para renomear os arquivos na pasta especificada
renomear_arquivos_com_sufixo(pasta, sufixos)
