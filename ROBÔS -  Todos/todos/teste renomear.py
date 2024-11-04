import os
import re
import pdfplumber

def extrair_cpf_cnpj(texto):
    # Padrões para CPF e CNPJ
    cpf_pattern = re.compile(r'\d{3}\.\d{3}\.\d{3}-\d{2}')
    cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
    
    # Encontrar todas as ocorrências usando findall
    cpfs = cpf_pattern.findall(texto)
    cnpjs = cnpj_pattern.findall(texto)
    
    # Combinar lista de CPFs e CNPJs
    documentos = cpfs + cnpjs
    
    # Debug: imprimir todas as ocorrências encontradas
    print(f"Todas as ocorrências encontradas: {documentos}")
    
    # Selecionar a segunda ocorrência se houver
    if len(documentos) >= 2:
        segundo_documento = documentos[1].replace('.', '').replace('/', '').replace('-', '')
        return segundo_documento

    return None

def ler_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = ""
            for page in pdf.pages:
                extracted_text = page.extract_text()
                if extracted_text:  # Certifique-se de que o texto não é None
                    texto += extracted_text + "\n"
            return texto
    except Exception as e:
        print(f"Erro ao ler PDF {caminho_pdf}: {e}")
        return ""

def renomear_pdfs(diretorio):
    for filename in os.listdir(diretorio):
        if filename.endswith('.pdf'):
            caminho_arquivo = os.path.join(diretorio, filename)
            texto_do_pdf = ler_pdf(caminho_arquivo)
            
            print(f"Texto extraído do '{filename}':\n{texto_do_pdf}\n")

            cpf_cnpj = extrair_cpf_cnpj(texto_do_pdf)
            if cpf_cnpj:
                novo_nome = os.path.join(diretorio, f"{os.path.splitext(filename)[0]}_{cpf_cnpj}_01.pdf")
                try:
                    os.rename(caminho_arquivo, novo_nome)
                    print(f"Arquivo renomeado para: {novo_nome}")
                except Exception as e:
                    print(f"Erro ao renomear '{filename}': {e}")
            else:
                print(f"Segunda ocorrência de CPF/CNPJ não encontrada no arquivo: {filename}")

# Defina o diretório onde estão os PDFs
diretorio_pdfs = r'S:/CAJ/Servidores/Murielle/TESTE'

# Chame a função para renomear os PDFs
renomear_pdfs(diretorio_pdfs)
