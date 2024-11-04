from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

# Definindo as variáveis com os dados de login (ALTERAR DADOS)
usuario = "USUÁRIO"
senha = "SENHA"

# Abrindo o navegador e acessando o site (ALTERAR DADOS)
driver = webdriver.Chrome()  # Utilize o driver apropriado para o seu navegador
driver.get("http://crago.cra21.com.br/crago/site/admin.php")
driver.implicitly_wait(10)  # Aguarda até 10 segundos para todos os elementos serem encontrados

# Localizando os campos de login
campo_usuario = WebDriverWait(driver, 5).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="login"]'))
)
campo_senha = WebDriverWait(driver, 5).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="senha"]'))
)

# Preenchendo os campos com as credenciais
campo_usuario.send_keys(usuario)
campo_senha.send_keys(senha)

# Localizando e clicando no botão de login
botao_enviar = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/div/div/div/div[2]/div/div[2]/div[4]/button/span'))
)
botao_enviar.click()

# Ícone das opções do menu
botao_menu = WebDriverWait(driver, 2).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/a'))
)
botao_menu.click()

# Clicar em Relatório
botao_relatorio1 = WebDriverWait(driver, 2).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[1]/div/ul/li[13]/a/i'))
)
botao_relatorio1.click()

# Clicar em relatório (aba)
botao_relatorio2 = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="submenu_menu_13"]/li/a/span'))
)
botao_relatorio2.click()

# Clicar opções status
botao_status = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/span/span[1]/span'))
)
botao_status.click()

# Selecionar status RETORNO
botao_retorno = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/span/span/span[2]/ul/li[3]'))
)
botao_retorno.click()

# Localizando e clicando no campo de seleção de data (inicial)
campo_data_inicial = WebDriverWait(driver, 5).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="dataInicial"]'))
)
campo_data_inicial.click()  # Abre o seletor de data

# Selecionar a data "HOJE" inicial
botao_data_inicial = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div[1]/table/tfoot/tr[1]/th'))
)
botao_data_inicial.click()

# Localizando e clicando no campo de seleção de data (final)
campo_data_final = WebDriverWait(driver, 5).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="dataFinal"]'))
)
campo_data_final.click()  # Abre o seletor de data

# Selecionar a data "HOJE" final
botao_data_final = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div[1]/table/tfoot/tr[1]/th'))
)
botao_data_final.click()

# Espera para garantir que tudo está carregado corretamente
time.sleep(2)

# Método alternativo para marcar checkbox: Arquivo Excel usando JavaScript
botao_checkbox_xpath = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="filtroTipoImpressaoRetorno"]/div/label[5]/span'))
)
botao_checkbox_xpath.click()

# Clicar em IMPRIMIR
botao_imprimir = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="imprimir-relatorio"]'))
)
botao_imprimir.click()

time.sleep(5)

# Fechar o navegador
driver.quit()

# Feito isso, o download do arquivo a ser analisado terá sido feito. Agora, é só alterar
# o caminho do arquivo a ser lido, no BOT 2.
