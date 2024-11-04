import os
import pandas as pd
import re

# Caminho do arquivo Excel e da pasta
arquivo_excel = r"S:/CAJ/Servidores/Murielle/X90J260724 Goiânia(1)- Usar para GERAR BOLETO E RENOMEAR.xlsx"
pasta = r"S:\CAJ\Servidores\Murielle\Remessa 26-07"

# Função para ler os dados das colunas CPF/CNPJ e Sequencial do arquivo Excel
def ler_planilha(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, dtype=str, header=0)  # Lê o arquivo Excel a partir da terceira linha
    df.columns = df.columns.str.strip()  # Remove espaços adicionais dos nomes das colunas
    print("Colunas disponíveis:", df.columns)
    
    colunas = ['CPFCNPJ', 'NUMEROTITULO', 'Sequencial']

    # Verifica se todas as colunas necessárias estão presentes
    for col in colunas:
        if col not in df.columns:
            print(f"A coluna '{col}' não está presente no arquivo.")
            return None  # Saia da função se alguma coluna não estiver presente

    df = df[colunas].dropna(subset=colunas)
    
    # Remover caracteres não numéricos da coluna CPF/CNPJ
    df['CPFCNPJ'] = df['CPFCNPJ'].apply(lambda x: re.sub(r'\D', '', x))
    
    # Formatar o campo CPF/CNPJ
    def formatar_cpf_cnpj(valor):
        num_digits = len(valor)
        if num_digits <= 11:
            return valor.zfill(11)  # CPF com 11 dígitos
        elif num_digits <= 14:
            return valor.zfill(14)  # CNPJ com 14 dígitos
        else:
            return valor

    df['CPFCNPJ'] = df['CPFCNPJ'].apply(formatar_cpf_cnpj)
    
    # Verificar se todos os valores foram formatados corretamente
    df['CPFCNPJ'] = df['CPFCNPJ'].apply(lambda x: x if len(x) in {11, 14} else None)
    if df['CPFCNPJ'].isnull().any():
        print("Erro de formatação nos seguintes valores:", df[df['CPFCNPJ'].isnull()])
    
    return df

# Função para renomear os arquivos adicionando os sufixos do CPF/CNPJ e Sequencial
def renomear_arquivos(pasta, dados):
    try:
        arquivos = os.listdir(pasta)
    except FileNotFoundError:
        print(f"Erro: O caminho da pasta '{pasta}' não foi encontrado.")
        return

    for nome_arquivo in arquivos:
        caminho_atual = os.path.join(pasta, nome_arquivo)
        
        if os.path.isfile(caminho_atual):
            nome_base, extensao = os.path.splitext(nome_arquivo)
            
            # Preserva o número guia do nome do arquivo
            partes = nome_base.split('_')
            numero_titulo = partes[0] if partes else nome_base
            
            # Procurar o registro na planilha baseado no número guia
            registro = dados[dados['NUMEROTITULO'] == numero_titulo]
            
            if not registro.empty:
                cpf_cnpj_formatado = registro.iloc[0]['CPFCNPJ']
                sequencial_formatado = registro.iloc[0]['Sequencial'] if 'Sequencial' in registro.columns else '01'
                
                # Novo nome do arquivo: NumeroGuia_CPF_CNPJ_Sequencial.extensão
                novo_nome_arquivo = f"{numero_titulo}_{cpf_cnpj_formatado}_{sequencial_formatado}{extensao}"
                novo_caminho = os.path.join(pasta, novo_nome_arquivo)
                
                try:
                    os.rename(caminho_atual, novo_caminho)
                    print(f"Arquivo renomeado: {nome_arquivo} -> {novo_nome_arquivo}")
                except Exception as e:
                    print(f"Erro ao renomear {nome_arquivo}: {e}")
            else:
                print(f"Registro não encontrado na planilha para {numero_titulo}")

# Lê as colunas CPF/CNPJ, Número Guia e Sequencial do arquivo Excel
dados = ler_planilha(arquivo_excel)

# Verifica se os dados foram lidos corretamente antes de prosseguir
if dados is not None:
    # Chama a função para renomear os arquivos na pasta especificada
    renomear_arquivos(pasta, dados)
