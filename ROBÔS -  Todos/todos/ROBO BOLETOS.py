import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import requests

# Configura o caminho para o chromedriver
chrome_driver_path = r'C:\Users\flcarneiro\.wdm\drivers\chromedriver\win64\125.0.6422.141\chromedriver.exe'
service = Service(chrome_driver_path)

# Inicializar o navegador
driver = webdriver.Chrome(service=service)

    # Abrir o site
driver.get("https://projudi.tjgo.jus.br/GerarBoleto?PaginaAtual=4")

    # Esperar a página carregar e garantir que os elementos estão presentes
wait = WebDriverWait(driver,5)

# Caminho do arquivo Excel
file_path = r'C:\Users\flcarneiro\Documents\protesto.xlsx'

# Nome da aba (sheet) no arquivo Excel
sheet_name = 'Conferência - GO'

# Lê o arquivo Excel, especificando a aba que contém os dados
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Limpando os nomes das colunas removendo espaços extras
df.columns = df.columns.str.strip()

# Verifica se a coluna "Número Guia" existe no DataFrame
coluna_nome_guia = 'Número Guia'
if coluna_nome_guia in df.columns:
    # Lista todos os valores da coluna "Número Guia"
    numeros_guia = df[coluna_nome_guia].tolist()
    
    # Imprime o primeiro número da coluna "Número Guia"
    if numeros_guia:
        primeiro_numero_guia = numeros_guia[0]
        print(f"Primeiro número da coluna '{coluna_nome_guia}': {primeiro_numero_guia}")
    else:
        print(f"A coluna '{coluna_nome_guia}' está vazia.")

try:
    # Preencher o campo numeroGuiaConsulta com o primeiro número da guia
    campo_guia = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="numeroGuiaConsulta"]'))
    )
    campo_guia.clear()
    campo_guia.send_keys(primeiro_numero_guia)
    
    # Clicar no botão com XPath //*[@id="divBotoesCentralizados"]/button[1]
    botao_buscar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="divBotoesCentralizados"]/button[1]'))
    )
    botao_buscar.click()
  # Aguarde alguns segundos para a ação ser concluída
    time.sleep(5)
    # Encontrar e preencher o campo 'Nome' com o valor da coluna 'Nome Parte'
    campo_nome_parte = driver.find_element(By.XPATH, '//*[@id="Nome"]')
    nome_parte = df['Nome Parte'].iloc[0]  
    campo_nome_parte.send_keys(nome_parte)
    print(f"Nome '{nome_parte}' inserido com sucesso no campo 'Nome' do site.")

    # Encontrar e preencher o campo 'Cpf' com o valor da coluna 'CPF/CNPJ'
    campo_cpf = driver.find_element(By.XPATH, '//*[@id="Cpf"]')
    cpf_cnpj = str(df['CPF/CNPJ'].iloc[0])  # Obtém o valor da coluna 'CPF/CNPJ'
    
    # Ajuste o formato do CPF/CNPJ, se necessário
    cpf_cnpj_formatado = "".join(filter(str.isdigit, str(cpf_cnpj)))  # Remove caracteres não numéricos
    
    campo_cpf.send_keys(cpf_cnpj_formatado)
    print(f"CPF/CNPJ '{cpf_cnpj_formatado}' inserido com sucesso no campo 'CPF' do site.")

    campo_logradouro = driver.find_element(By.XPATH, '//*[@id="Logradouro"]')
    endereco_formatado = df['Endereço Formatado'].iloc[0]  # Obtém o valor da coluna 'Endereço Formatado'
    
    campo_logradouro.send_keys(endereco_formatado)
    print(f"Endereço '{endereco_formatado}' inserido com sucesso no campo 'Logradouro' do site.")

    campo_bairro = driver.find_element(By.XPATH, '//*[@id="Bairro"]')
    bairro = df['Bairro'].iloc[0]  # Obtém o valor da coluna 'Bairro'
    
    campo_bairro.send_keys(bairro)
    print(f"Bairro '{bairro}' inserido com sucesso no campo 'Bairro' do site.")

    campo_cidade = driver.find_element(By.XPATH, '//*[@id="Cidade"]')
    cidade = df['Cidade'].iloc[0]  # Obtém o valor da coluna 'Cidade'
    
    campo_cidade.send_keys(cidade)
    print(f"Cidade '{cidade}' inserida com sucesso no campo 'Cidade' do site.")

    campo_uf = driver.find_element(By.XPATH, '//*[@id="Uf"]')
    uf = df['UF'].iloc[0]  # Obtém o valor da coluna 'UF'
    
    campo_uf.send_keys(uf)
    print(f"UF '{uf}' inserida com sucesso no campo 'UF' do site.")

    campo_cep = driver.find_element(By.XPATH, '//*[@id="Cep"]')
    cep = str(df['CEP'].iloc[0])  # Obtém o valor da coluna 'CEP'
    
    campo_cep.send_keys(cep)
    print(f"CEP '{cep}' inserido com sucesso no campo 'CEP' do site.")

    # Encontrar e clicar no botão
    botao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divBotoesCentralizados"]/button[1]')))
    botao.click()
    
    print("Clicou com sucesso no botão.")

except Exception as e:
    print(f"Erro ao clicar no botão: {e}")


    botao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divBotoesCentralizados"]/button')))
    botao.click()
    
    print("Clicou com sucesso no botão.")

except Exception as e:
    print(f"Erro ao clicar no botão: {e}")

try:
    # Clicar em "Emitir e Imprimir" no campo //*[@id="imgEmitirGuiaImprimirPDF"]
    botao_emissao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="imgEmitirGuiaImprimirPDF"]')))
    botao_emissao.click()
    
    print("Clicou em 'Emitir e Imprimir' no campo 'imgEmitirGuiaImprimirPDF'.")
    time.sleep(5)
    
    botao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divBotoesCentralizados"]/button')))
    botao.click()
    
    print("Clicou em 'Emitir e Imprimir' no campo 'Emitir e Imprimir'.")
    time.sleep(5)

except Exception as e:
    print(f"Erro ao clicar no botão 'Emitir e Imprimir': {e}")
