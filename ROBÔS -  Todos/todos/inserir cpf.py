import os
import pandas as pd
import re

# Caminho do arquivo Excel e da pasta
arquivo_excel = r"S:/CAJ/Servidores/Murielle/Demanda GISELE -  Setembro  - SEM DUPLICATAS.xlsx"
pasta = r"S:\CAJ\Servidores\Murielle\Demanda CRA - Setembro"

# Função para ler os dados das colunas CPF/CNPJ e Sequencial do arquivo Excel
def ler_planilha(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, dtype=str)  # Lê o arquivo Excel do início
    df.columns = [str(col).strip() for col in df.columns]  # Remover espaços extras dos nomes das colunas
    print("Colunas disponíveis:", df.columns)  # Imprimir nomes das colunas para depuração
    
    # Nome das colunas a considerar
    colunas = ['CPF/CNPJ', 'Número Guia', 'Sequencial']
    df = df[colunas].dropna(subset=colunas)
    
    # Remover caracteres não numéricos das colunas
    df['CPF/CNPJ'] = df['CPF/CNPJ'].apply(lambda x: re.sub(r'\D', '', x))
    
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
def renomear_arquivos(pasta, dados):
    try:
        arquivos = os.listdir(pasta)
    except FileNotFoundError:
        print(f"Erro: O caminho da pasta '{pasta}' não foi encontrado. Verifique se o caminho está correto.")
        return
    
    for nome_arquivo in arquivos:
        caminho_atual = os.path.join(pasta, nome_arquivo)
        
        if os.path.isfile(caminho_atual):
            nome_base, extensao = os.path.splitext(nome_arquivo)
            
            # Preserva o número guia do nome do arquivo
            partes = nome_base.split('_')
            numero_guia = partes[0] if partes else nome_base
            
            # Procurar o registro na planilha baseado no número guia
            registro = dados[dados['Número Guia'] == numero_guia]
            
            if not registro.empty:
                cpf_cnpj_formatado = registro.iloc[0]['CPF/CNPJ']
                sequencial_formatado = '01'
                
                # Novo nome do arquivo: NumeroGuia_CPF_CNPJ_01.extensao
                novo_nome_arquivo = f"{numero_guia}_{cpf_cnpj_formatado}_{sequencial_formatado}{extensao}"
                novo_caminho = os.path.join(pasta, novo_nome_arquivo)
                
                os.rename(caminho_atual, novo_caminho)
                print(f"Arquivo renomeado: {nome_arquivo} -> {novo_nome_arquivo}")
            else:
                print(f"Registro não encontrado na planilha para {numero_guia}")

# Lê as colunas CPF/CNPJ, Número Guia e Sequencial do arquivo Excel
dados = ler_planilha(arquivo_excel)

# Chama a função para renomear os arquivos na pasta especificada
renomear_arquivos(pasta, dados)
