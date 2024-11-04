import os
import re
from PyPDF2 import PdfReader

def extrair_cpf_cnpj_do_pagador(texto):
    # Padrões para CPF e CNPJ
    cpf_pattern = re.compile(r'\d{3}\.\d{3}\.\d{3}-\d{2}')
    cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
    
    # Encontrar todas as ocorrências de CPF e CNPJ
    cpfs = cpf_pattern.findall(texto)
    cnpjs = cnpj_pattern.findall(texto)

    # Unir os resultados em uma lista
    todos_os_numeros = cpfs + cnpjs

    # Verificar se há pelo menos dois números encontrados
    if len(todos_os_numeros) >= 2:
        # Retorna o segundo número encontrado (removendo pontuações)
        return todos_os_numeros[1].replace('.', '').replace('-', '').replace('/', '')

    return None

def ler_pdf(caminho_pdf):
    # Lê o conteúdo do PDF
    try:
        with open(caminho_pdf, 'rb') as file:
            reader = PdfReader(file)
            texto = ""
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                texto += page.extract_text() or ""  # Garantir que nenhum texto nulo cause problemas
            return texto
    except Exception as e:
        print(f"Erro ao ler o arquivo {caminho_pdf}: {e}")
        return None

def renomear_pdfs(diretorio):
    for filename in os.listdir(diretorio):
        if filename.endswith('.pdf'):
            caminho_arquivo = os.path.join(diretorio, filename)
            texto_do_pdf = ler_pdf(caminho_arquivo)

            if texto_do_pdf:  # Verifique se o texto foi lido corretamente
                cpf_cnpj_pagador = extrair_cpf_cnpj_do_pagador(texto_do_pdf)
                if cpf_cnpj_pagador:
                    novo_nome = os.path.join(diretorio, f"{os.path.splitext(filename)[0]}_{cpf_cnpj_pagador}_01.pdf")
                    os.rename(caminho_arquivo, novo_nome)
                    print(f"Arquivo renomeado para: {novo_nome}")
                else:
                    print(f"Não foi possível encontrar o segundo CPF/CNPJ (do pagador) no arquivo: {filename}")
            else:
                print(f"Texto não pôde ser lido do arquivo: {filename}")

# Defina o diretório onde estão os PDFs
diretorio_pdfs = r'S:\CAJ\Servidores\Murielle\boletos\deram erro'

# Chame a função para renomear os PDFs
renomear_pdfs(diretorio_pdfs)
