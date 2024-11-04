import os
import shutil
import pandas as pd
from getpass import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import logging

# Função para configurar o logging
def setup_logging():
    logging.basicConfig(
        filename='processamento.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

def log_and_print(message):
    logging.info(message)
    print(message)

# Função para abrir o site e realizar o login
def abrir_site_e_login(url, username, password):
    try:
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service)
        driver.get(url)
        user_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "login"))
        )
        pass_field = driver.find_element(By.ID, "senha")
        user_field.send_keys(username)
        pass_field.send_keys(password)
        pass_field.send_keys(Keys.RETURN)
        return driver
    except Exception as e:
        log_and_print(f"Erro ao abrir o site e realizar login: {e}")
        return None

# Função para copiar o número do processo
def copiar_numero_processo(driver, numero_processo):
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for iframe in iframes:
            driver.switch_to.frame(iframe)
            try:
                processo_input = driver.find_element(By.ID, "ProcessoNumero")
                if processo_input:
                    break
            except:
                driver.switch_to.default_content()
                continue

        processo_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        processo_input.clear()
        processo_input.send_keys(numero_processo)
        log_and_print(f"Colou o número do processo: {numero_processo}")
    except Exception as e:
        log_and_print(f"Erro ao colar o número do processo: {e}")

# Função para clicar no botão "Consultar"
def clicar_botao_consultar(driver):
    try:
        botao_consultar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/div/input"))
        )
        botao_consultar.click()
        log_and_print("Botão 'Consultar' clicado.")
    except Exception as e:
        log_and_print(f"Erro ao clicar no botão 'Consultar': {e}")

# Função para localizar o ID do processo e clicar em "Editar Débito"
def localizar_id_e_clicar_editar(driver, id_processo):
    try:
        tabela = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                if coluna.text.strip() == str(id_processo):
                    log_and_print(f"ID {id_processo} encontrado na tabela.")
                    botao_editar = linha.find_element(By.XPATH, ".//input[@type='image' and contains(@title, 'Editar Débito')]")
                    botao_editar.click()
                    log_and_print(f"Clicou em 'Editar Débito' para o ID: {id_processo}")
                    return True
        log_and_print(f"ID do processo {id_processo} não encontrado na tabela.")
        return False
    except Exception as e:
        log_and_print(f"Erro ao localizar o ID do processo e clicar em 'Editar Débito': {e}")
        return False

# Função para alterar o status do débito
def alterar_status_debito(driver, novo_status):
    try:
        select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "Id_ProcessoDebitoStatus"))
        )
        select = Select(select_element)
        select.select_by_visible_text(novo_status)
        log_and_print(f"Alterou o status do débito para: {novo_status}")
    except Exception as e:
        log_and_print(f"Erro ao alterar o status do débito: {e}")

# Função para clicar no botão "Salvar"
def clicar_botao_salvar(driver):
    try:
        botao_salvar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divPortaBotoes"]/button[1]'))
        )
        botao_salvar.click()
        log_and_print("Botão 'Salvar' clicado.")
    except Exception as e:
        log_and_print(f"Erro ao clicar no botão 'Salvar': {e}")

# Função para clicar no botão "Confirmar Salvar"
def clicar_botao_confirmar_salvar(driver):
    try:
        botao_confirmar_salvar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divConfirmarSalvar"]/button'))
        )
        botao_confirmar_salvar.click()
        log_and_print("Botão 'Confirmar Salvar' clicado.")
    except Exception as e:
        log_and_print(f"Erro ao clicar no botão 'Confirmar Salvar': {e}")

# Função para clicar no botão "Ok"
def clicar_botao_ok(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[8]/div[3]/div/button'))
        )
        botao_ok = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[3]/div/button'))
        )
        botao_ok.click()
        log_and_print("Botão 'Ok' clicado.")
    except Exception as e:
        log_and_print(f"Erro ao clicar no botão 'Ok': {e}")

# Função para clicar no botão "Escolher Novo Processo" pelo atributo title
def clicar_botao_escolher_novo_processo(driver):
    try:
        botao_novo_processo = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@title="Escolhe novo processo"]'))
        )
        botao_novo_processo.click()
        log_and_print("Botão 'Escolher Novo Processo' clicado.")
    except Exception as e:
        log_and_print(f"Erro ao clicar no botão 'Escolher Novo Processo': {e}")

# Atualizar o caminho da pasta
folder_path = r'S:\CAJ\Cobrança - Judicial e Administrativa\Custas Judiciais - Finais\Baixa de averbação - pagos - status novo\Aguardando processamento'
processed_folder_path = r'S:\CAJ\Cobrança - Judicial e Administrativa\Custas Judiciais - Finais\Baixa de averbação - pagos - status novo\Processados'

# Configuração do logging
setup_logging()

# Encontrar o arquivo Excel na pasta
try:
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not excel_files:
        log_and_print(f"Erro: Nenhum arquivo .xlsx encontrado na pasta especificada: {folder_path}")
        exit(1)
    file_path = os.path.join(folder_path, excel_files[0])
except Exception as e:
    log_and_print(f"Erro ao procurar o arquivo na pasta especificada: {folder_path}\n{e}")
    exit(1)

required_columns = ['ID', 'Número Processo', 'Status Guia']
try:
    df = pd.read_excel(file_path, sheet_name='relatorio-processo-parte-debito')
    if not all(column in df.columns for column in required_columns):
        log_and_print(f"Erro: O arquivo excel não contém todas as colunas necessárias: {required_columns}")
        exit(1)
except FileNotFoundError as e:
    log_and_print(f"Erro: Arquivo não encontrado no caminho especificado: {file_path}\n{e}")
    exit(1)
except Exception as e:
    log_and_print(f"Erro ao ler o arquivo Excel: {e}")
    exit(1)

log_and_print(df.columns)
log_and_print(df.head())

# Adicionar a coluna "Resultado do Processamento"
df["Resultado do Processamento"] = ""

# Credenciais e URL
url = "https://projudi.tjgo.jus.br/"
username = input("Digite seu usuário: ")
password = getpass("Digite sua senha: ")

if not username or not password:
    log_and_print("As credenciais são obrigatórias.")
    exit(1)

driver = abrir_site_e_login(url, username, password)
if driver is None:
    log_and_print("Falha ao abrir o site e realizar o login. Encerrando o script.")
    exit(1)

try:
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Analista Financeiro"))
        )
        driver.find_element(By.LINK_TEXT, "Analista Financeiro").click()
    except Exception as e:
        log_and_print("Elemento 'Analista Financeiro' não encontrado, prosseguindo para 'Cadastros'.")

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Cadastros"))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Financeiros"))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Débitos - Diretoria Financeira"))
    ).click()

    for index, row in df.iterrows():
        if pd.isna(row['ID']):
            break

        numero_processo = row['Número Processo']
        processo_id = row['ID']
        status_guia = row['Status Guia']

        copiar_numero_processo(driver, numero_processo)
        clicar_botao_consultar(driver)

        sucesso = localizar_id_e_clicar_editar(driver, processo_id)
        if not sucesso:
            df.at[index, "Resultado do Processamento"] = "ID do processo não encontrado no sistema"
            clicar_botao_escolher_novo_processo(driver)
            continue

        try:
            status_debito_select = Select(driver.find_element(By.ID, "Id_ProcessoDebitoStatus"))
            status_debito = status_debito_select.first_selected_option.text
            if status_debito in ["Novo", "Aguardando Pagamento"]:
                if status_guia in ["PAGO", "PAGO APÓS O VENCIMENTO"]:
                    alterar_status_debito(driver, "Pago")
                    clicar_botao_salvar(driver)
                    clicar_botao_confirmar_salvar(driver)
                    clicar_botao_ok(driver)
                    df.at[index, "Resultado do Processamento"] = "Status alterado para Pago"
                else:
                    df.at[index, "Resultado do Processamento"] = "Status da guia não é 'PAGO' ou 'PAGO APÓS O VENCIMENTO'"
            else:
                df.at[index, "Resultado do Processamento"] = "Status do débito não é 'Novo' ou 'Aguardando Pagamento'"
        except Exception as e:
            df.at[index, "Resultado do Processamento"] = f"Erro ao verificar ou alterar o status do débito: {e}"

        clicar_botao_escolher_novo_processo(driver)

finally:
    output_file_name = os.path.splitext(os.path.basename(file_path))[0] + " - Resultado do Processamento.xlsx"
    output_file_path = os.path.join(processed_folder_path, output_file_name)
    df.to_excel(output_file_path, index=False)
    log_and_print(f"Resultados do processamento salvos em: {output_file_path}")

    try:
        shutil.move(file_path, processed_folder_path)
        log_and_print(f"Arquivo movido para a pasta: {processed_folder_path}")
    except Exception as e:
        log_and_print(f"Erro ao mover o arquivo para a pasta 'Processados': {e}")

    driver.quit()
