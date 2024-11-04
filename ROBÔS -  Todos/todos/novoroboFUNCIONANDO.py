import os
import pandas as pd
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


# Configura o caminho para o chromedriver
chrome_driver_path = r'C:\Users\flcarneiro\.wdm\drivers\chromedriver\win64\125.0.6422.141\chromedriver.exe'
service = Service(chrome_driver_path)

# Inicializar o navegador
driver = webdriver.Chrome(service=service)

    # Abrir o site
driver.get("https://depositojudicial.caixa.gov.br/sigsj_internet/login.xhtml")

    # Esperar a página carregar e garantir que os elementos estão presentes
wait = WebDriverWait(driver, 20)

    # Tentar encontrar e preencher o campo de usuário
usuario_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'usuario') or contains(@name, 'usuario') or contains(@id, 'username') or contains(@name, 'username')]")))
usuario_field.clear()
usuario_field.send_keys("FINANCEIRA2024")

    # Tentar encontrar e preencher o campo de senha
senha_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'senha') or contains(@name, 'senha') or contains(@id, 'password') or contains(@name, 'password')]")))
senha_field.clear()
senha_field.send_keys("caj2024")

    # Enviar o formulário de login
senha_field.submit()

def copiar_dados_planilha(caminho_arquivo):
    """
    Lê a planilha Excel e retorna um DataFrame com os dados.
    """
    try:
        # Verificar extensão do arquivo
        if caminho_arquivo.endswith('.xlsx'):
            df = pd.read_excel(caminho_arquivo, engine='openpyxl')
        elif caminho_arquivo.endswith('.xls'):
            df = pd.read_excel(caminho_arquivo, engine='xlrd')
        else:
            raise ValueError("Formato de arquivo não suportado.")
        
        print("Colunas encontradas na planilha:", df.columns.tolist())
        return df
    except Exception as e:
        print(f"Erro ao copiar os dados da planilha: {e}")
        return None



def configurar_driver():
    """
    Configura o driver do Chrome usando webdriver_manager.
    """
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    return driver

# Ler os dados da planilha "caixa" na aba
def ler_dados_planilha(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name='05-2024', header=0)
        return df['Agencia'].tolist()
    except Exception as e:
        print(f"Erro ao ler os dados da planilha: {e}")
        return None
caminho_arquivo = r'C:\Users\flcarneiro\Desktop\SCRIPTS\caixa1.xlsx'
nome_aba = '05-2024'
linha_nomes_colunas =0  

# Carregar o arquivo Excel
df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba, header=linha_nomes_colunas)
print(df.columns)
# Copiar os dados da coluna "Agencia" para o campo "Agência"
df['Agencia'] = df['Agencia']
# Obter o primeiro número da coluna "Agencia"
primeiro_numero_agencia = df['Agencia'].iloc[0]
# Encontrar o índice da coluna "CONTA" e "DV"
indice_coluna_conta = df.columns.get_loc("CONTA")
indice_coluna_dv = df.columns.get_loc("DV")


# Obter os valores das primeiras linhas das colunas "CONTA" e "DV"
primeira_conta = df.iloc[0, indice_coluna_conta]
primeiro_dv = df.iloc[0, indice_coluna_dv]

# Local específico para colar o número, neste caso, assumindo uma variável "local_paste"
local_paste = primeiro_numero_agencia

print(f'O primeiro número da coluna "Agencia" é: {primeiro_numero_agencia}')
print(f'Você pode colá-lo no local específico: {local_paste}')

# Encontrar o campo "Agência" e preenchê-lo com o número encontrado
campo_agencia = driver.find_element(By.XPATH, '//*[@id="j_id49:filtroView:j_id50:agencia"]')
campo_agencia.send_keys(str(primeiro_numero_agencia))
print(f'Preencheu o campo Agência com: {primeiro_numero_agencia}')
# Encontrar e preencher o campo "Operação"
campo_operacao = driver.find_element(By.XPATH, '//*[@id="j_id49:filtroView:j_id50:operacaoTrue"]')
campo_operacao.send_keys("040 - Depósitos Judiciais da Justiça Estadual")
print('Preencheu o campo Operação com: 040 - Depósitos Judiciais da Justiça Estadual')
# Encontrar e preencher o campo "CONTA" com o valor da primeira linha da coluna "CONTA"
campo_conta = driver.find_element(By.XPATH, '//*[@id="j_id49:filtroView:j_id50:conta"]')
campo_conta.send_keys(str(primeira_conta))
print(f'Preencheu o campo Conta com: {primeira_conta}')

# Encontrar e preencher o campo "DV" com o valor da primeira linha da coluna "DV"
campo_dv = driver.find_element(By.XPATH, '//*[@id="j_id49:filtroView:j_id50:dv"]')
campo_dv.send_keys(str(primeiro_dv))
print(f'Preencheu o campo DV com: {primeiro_dv}')
# Encontrar e clicar no botão "Consultar"
botao_consultar = driver.find_element(By.XPATH, '//*[@id="j_id49:filtroView:j_id50:btConsultaConta"]')
botao_consultar.click()
print('Clicou no botão Consultar')

try:
    # Encontrar e clicar na imagem com XPath
    imagem_extrato = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="j_id49:filtroView:j_id215:listaView:j_id240:0:j_id305:0:extrato"]/img'))
    )
    imagem_extrato.click()
    print('Clicou na imagem de Extrato')
except Exception as e:
    print(f"Erro ao clicar na imagem de 'Extrato': {e}")

try:
    # Encontrar e clicar no botão "Período" com XPath
    botao_periodo = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="j_id49:filtroView:j_id1458:icPadraoPesquisa"]/tbody/tr[2]/td/span/span/img'))
    )
    botao_periodo.click()
    print('Clicou no botão Período')
except Exception as e:
    print(f"Erro ao clicar no botao 'período': {e}")

# Definir a data desejada
data_desejada = "31/05/2024"

# Encontrar o campo de input pelo XPath
campo_data = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="j_id49:filtroView:j_id1458:dtInicialInputDate"]'))
)

campo_data.clear()  # Limpar o campo antes de preencher
campo_data.send_keys(Keys.CONTROL + "a")  # Selecionar todo o texto no campo
campo_data.send_keys("31/05/2024")  # Inserir a nova data


# Encontrar o campo de input pelo XPath novamente
campo_data = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="j_id49:filtroView:j_id1458:dtFinalInputDate"]'))
)
campo_data.clear()  # Limpar o campo antes de preencher
campo_data.send_keys(Keys.CONTROL + "a")  # Selecionar todo o texto no campo
campo_data.send_keys("31/05/2024")  # Inserir a nova data
campo_data.send_keys(Keys.ENTER)  # Pressionar a tecla ENTER

try:
    campo_lista = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="j_id49:filtroView:j_id1458:listaView:lista:2:j_id1540"]'))
    )
    texto_campo = campo_lista.text
    print("Conteúdo do campo:")
    print(texto_campo)
except NoSuchElementException as e:
    print(f"Elemento não pôde ser encontrado: {e}")

# Adicionar a informação à coluna "2024-05-31 00:00:00" em algum lugar, como uma planilha ou banco de dados
caminho_arquivo_excel = "C:/Users/flcarneiro/Desktop/SCRIPTS/caixa1.xlsx"
aba = "05-2024"
coluna = "2024-05-31 00:00:00"
linha = 0  # Note que 'loc' baseia-se em rótulos de linha, se você quer a segunda linha, use índice 1

# Carregar o arquivo Excel e obter o DataFrame
df = pd.read_excel(caminho_arquivo_excel, sheet_name=aba)

# Verifica se a coluna existe, caso contrário, cria a coluna
if coluna not in df.columns:
    df[coluna] = None

# Adicionar a informação na linha especificada e coluna correspondente
try:
    df.loc[linha, coluna] = texto_campo
    print(f"Número '{texto_campo}' inserido com sucesso na linha {linha + 1}, coluna {coluna}.")
except Exception as e:
    print(f"Erro ao inserir o número na planilha: {e}")

# Salvar as mudanças na planilha
df.to_excel(caminho_arquivo_excel, sheet_name=aba, index=False)
print(f'Dados coletados e salvos na planilha na aba "{aba}" na coluna "{coluna}": {texto_campo}')
