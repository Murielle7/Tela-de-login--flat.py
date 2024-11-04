import PyPDF2
import re
import os

def extrair_informacoes_pagador(arquivo_pdf):
    """Extrai as informações do pagador de um arquivo PDF.

    Args:
        arquivo_pdf (str): Caminho completo do arquivo PDF.

    Returns:
        dict: Dicionário contendo as informações do pagador, ou None se não encontradas.
    """

    with open(arquivo_pdf, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        page_obj = pdf_reader.pages[0]  # Assumindo que as informações estão na primeira página
        texto = page_obj.extract_text()

        # Ajustar as expressões regulares de acordo com os seus PDFs
        padrao_nome = r"Nome do Pagador:\s*(.*)"
        padrao_cpf_cnpj = r"CPF/CNPJ:\s*(\d{11}|\d{14})"

        match_nome = re.search(padrao_nome, texto, re.IGNORECASE)
        match_cpf_cnpj = re.search(padrao_cpf_cnpj, texto)

        if match_nome and match_cpf_cnpj:
            return {
                'nome': match_nome.group(1).strip(),
                'cpf_cnpj': match_cpf_cnpj.group(1).strip()
            }
        else:
            return None

# Pasta com os arquivos
pasta_arquivos = "S:\CAJ\Servidores\Murielle\boletos\deram erro"

# Processa cada arquivo PDF na pasta
for arquivo in os.listdir(pasta_arquivos):
    if arquivo.endswith(".pdf"):
        caminho_completo = os.path.join(pasta_arquivos, arquivo)
        informacoes = extrair_informacoes_pagador(caminho_completo)

        if informacoes:
            print(f"Arquivo: {arquivo}")
            print(informacoes)
        else:
            print(f"Informações do pagador não encontradas em {arquivo}")
