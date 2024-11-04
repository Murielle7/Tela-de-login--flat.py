import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Defina o caminho do arquivo (ALTERAR CASO NECESSÁRIO)
caminho_arquivo = r's:\CAJ\Servidores\Murielle\Demanda GISELE -  Setembro  - SEM DUPLICATAS.xlsx'

# Automatizar interação com o navegador usando Selenium
driver = webdriver.Chrome()  # Certifique-se de ter o ChromeDriver configurado e acessível
driver.get('https://projudi.tjgo.jus.br/')

# Credenciais de login
usuario = "70273080105"
senha = "235791Mu@"

# Realizar login
campo_usuario = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, 'login')))
campo_senha = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, 'senha')))
campo_usuario.send_keys(usuario)
campo_senha.send_keys(senha)

botao_entrar = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"]')))
botao_entrar.click()

# Função para preencher o número do processo sem recarregar a página
def preencher_numero_processo(nosso_numero):
    campo_processo_numero = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, 'ProcessoNumero')))
    campo_processo_numero.clear()
    campo_processo_numero.send_keys(nosso_numero)

try:
    # Leia a planilha 'Boletos - Protestadas- automaça' do arquivo Excel
    df = pd.read_excel(caminho_arquivo, sheet_name='Envio CRA', engine='openpyxl', header=0)
    print("Planilha 'Envio CRA' lida com sucesso a partir da linha 1!")

    for index, row in df.iterrows():
        if row['Status'] == 'Protestado':
            nosso_numero = str(row['NOSSONUMERO'])
            preencher_numero_processo(nosso_numero)
            
        # Acessando a seção Cadastro
        cadastro_xpath = '//*[@id="2"]'
        cadastro_element = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, cadastro_xpath))
        )
        cadastro_element.click()

        # Acessando a subseção Financeira
        financeira_xpath = '//*[@id="2m1"]'
        financeira_element = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, financeira_xpath))
        )
        financeira_element.click()

        # Acessando a sub-subseção Débitos – Diretoria Financeira
        debitos_xpath = '//*[@id="2m1m0"]'
        debitos_element = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, debitos_xpath))
        )
        debitos_element.click()

        # Preencher o campo Número processo !!!!! ATENÇÃO !!!!!
        campo_Processo_Numero = driver.find_element(By.XPATH, '//*[@id="ProcessoNumero"]')  # Localizar o elemento pelo xpath
        campo_Processo_Numero.clear()  # Limpar o campo antes de preencher com um novo número de guia
        campo_Processo_Numero.send_keys(nosso_numero)
        
        # Encontrar e clicar no botão "Consultar"
        botao_consultar = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btnConsultar"]'))
        )
        botao_consultar.click()

        # Clicar no botão 'Editar'
        botao_editar = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[2]/fieldset[2]/table/tbody/tr/td[11]/input '))  # Verifique se o XPath está correto
        )
        botao_editar.click()

        # Localizando e clicando no campo de seleção do Status
        campo_status = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="Id_ProcessoDebitoStatus"]'))  # Atualize o XPath conforme necessário
        )
        campo_status.click()

        # Selecionar o status ‘Protestado’
        status_protestado = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[2]/fieldset[1]/select/option[18]'))  # Atualize o XPath conforme necessário
        )
        status_protestado.click()

        # Clicar no botão ‘Salvar’
        botao_salvar = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[1]/button[1]/i'))  # Atualize o XPath conforme necessário
        )
        botao_salvar.click()

except Exception as e:
    print(f"Ocorreu um erro: {str(e)}")
finally:
    driver.quit()
