import pandas as pd
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.firefox import GeckoDriverManager
from getpass import getpass

def copiar_dados_planilha(origem, destino):
    try:
        df = pd.read_excel(origem, sheet_name='Planilha1', header=0)
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None
    
    colunas_necessarias = ['Número Processo', 'Número Guia', 'Status Débito', 'Nome Parte']  # Adiciona 'Nome Parte'
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print("Erro: uma ou mais colunas não foram encontradas na planilha 'Planilha1'.")
        return None
    
    df_filtrado = df[colunas_necessarias].copy()
    df_filtrado['Alteração de Status'] = ""  # Adiciona a nova coluna para status

    try:
        df_filtrado.to_excel(destino, index=False)
        print(f"Dados salvos com sucesso em: {destino}")
    except Exception as e:
        print(f"Erro ao salvar a nova planilha: {e}")

    return df_filtrado

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
        print(f"Erro ao abrir o site e realizar login: {e}")
        return None

def clicar_analista_financeiro(driver):
    try:
        analista_link = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Analista Financeiro')]"))
        )
        analista_link.click()
        print("Clicou em 'Analista Financeiro'.")
    except Exception as e:
        print(f"Erro ao clicar em 'Analista Financeiro': {e}")

def clicar_menu_cadastros(driver):
    try:
        cadastros_menu = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "2"))
        )
        cadastros_menu.click()
        print("Clicou no menu 'Cadastros'.")
        
        financeiros_submenu = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))
        )
        financeiros_submenu.click()
        print("Clicou no submenu 'Financeiros'.")

        debitos_submenu = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))
        )
        debitos_submenu.click()
        print("Clicou no submenu 'Débitos - Diretoria Financeira'.")
        
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

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
        print(f"Colou o número do processo: {numero_processo}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

# Função para clicar no botão "Consultar"
def clicar_botao_consultar(driver):
    try:
        botao_consultar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/div/input"))
        )
        botao_consultar.click()
        print("Botão 'Consultar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Consultar': {e}")

# Função para localizar o ID do processo e clicar em "Editar Débito"
def localizar_guia_e_clicar_editar(driver, numero_guia):
    try:
        # Garantir que cpf_cnpj é uma string
        numero_guia = str(numero_guia)
        
        tabela = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                parte_text = str(coluna.text).strip().replace(".", "").replace("/", "").replace("-", "")
                if guia_text == v.replace(".", "").replace("/", "").replace("-", ""):
                    print(f"CPF/CNPJ { numero_guia } encontrado na tabela.")
                    botao_editar = linha.find_element(By.XPATH, ".//input[@type='image' and contains(@title, 'Editar Débito')]")
                    botao_editar.click()
                    time.sleep(2)
                    print(f"Clicou em 'Editar Débito' para o CPF/CNPJ: { numero_guia }")
                    return True
        print(f"CPF/CNPJ { numero_guia } não encontrado na tabela.")
        return False
    except Exception as e:
        print(f"Erro ao localizar o CPF/CNPJ e clicar em 'Editar Débito': {e}")
        return False

# Função para alterar o status do débito
def alterar_status_debito(driver, novo_status):
    try:
              select_element = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.ID, "Id_ProcessoDebitoStatus"))
        )
              opcao_em_analise = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/form/div[2]/fieldset[1]/select/option[4]'))
                )
              opcao_em_analise.click()
              print("Selecionou a opção 'APTO PARA ENVIO AO PROTESTO'")
         
              time.sleep(2)
    except Exception as e:
         print(f"Erro ao inserir os dados no campo e realizar as ações: {e}")

# Função para clicar no botão "Salvar"
def clicar_botao_salvar(driver):
    try:
        botao_salvar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divPortaBotoes"]/button[1]'))
        )
        botao_salvar.click()
        print("Botão 'Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Salvar': {e}")

# Função para clicar no botão "Confirmar Salvar"
def clicar_botao_confirmar_salvar(driver):
    try:
        botao_confirmar_salvar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divConfirmarSalvar"]/button'))
        )
        botao_confirmar_salvar.click()
        print("Botão 'Confirmar Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Confirmar Salvar': {e}")

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
        print("Botão 'Ok' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Ok': {e}")

# Função para clicar no botão "Escolher Novo Processo" pelo atributo title
def clicar_botao_escolher_novo_processo(driver):
    try:
        botao_novo_processo = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, ' /html/body/div[1]/form/div[1]/button[2]/i '))
        )
        botao_novo_processo.click()
        print("Botão 'Escolher Novo Processo' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Escolher Novo Processo': {e}")

# Atualizar o caminho da pasta
def main():
    caminho_arquivo = r'S:/CAJ/Cobrança - Judicial e Administrativa/Custas Judiciais - Finais/Processamento Retorno - CRA/Relatório BI.xlsx'
    caminho_arquivo_destino = os.path.join(os.path.expanduser('~'), 'Desktop', 'Teste Python 123.xlsx')
    
    df_filtrado = copiar_dados_planilha(caminho_arquivo, caminho_arquivo_destino)
    if df_filtrado is None:
        return

# Credenciais e URL
url = "https://projudi.tjgo.jus.br/"
username = input("Digite seu usuário: ")
password = getpass("Digite sua senha: ")

driver = abrir_site_e_login(url, username, password)
if driver is None:
    print("Falha ao abrir o site e realizar o login. Encerrando o script.")
    exit(1)

    clicar_analista_financeiro(driver)
    clicar_menu_cadastros(driver)

    for index, row in df_filtrado.iterrows():
        numero_processo = row['Número Processo']
        processo_id = row['Número Guia']
        status_guia = row['Status']
        
        try:
            copiar_numero_processo(driver, numero_processo)
            clicar_botao_consultar(driver)

            sucesso = localizar_guia_e_clicar_editar(driver, numero_guia)
            if not sucesso:
                df.at[index, "Resultado do Processamento"] = "ID do processo não encontrado no sistema"
                clicar_botao_escolher_novo_processo(driver)
                continue

            try:
                status_debito_select = Select(driver.find_element(By.ID, "/html/body/div[1]/form/div[2]/fieldset[1]/select/option[19]"))
                status_debito = status_debito_select.first_selected_option.text
                if status_debito in ["Apto para envio ao protesto", "Em Análise pelo Financeiro"]:
                    if status_guia in ["AGUARDANDO PAGAMENTO"]:
                        alterar_status_debito(driver, "Protestado")
                        clicar_botao_salvar(driver)
                        clicar_botao_confirmar_salvar(driver)
                        clicar_botao_ok(driver)
                        df.at[index, "Resultado do Processamento"] = "Status alterado para Pago"
                    else:
                        df.at[index, "Resultado do Processamento"] = "Status da guia não é 'PAGO' ou 'PAGO APÓS O VENCIMENTO'"
                else:
                    df.at[index, "Resultado do Processamento"] = "Status do débito não é 'Novo' ou 'Aguardando Pagamento'"
            except Exception as e:
                df.at[index, "Resultado do Processamento"] = f"Erro ao verificar ou alterar o status do débito: {str(e)}"

            clicar_botao_escolher_novo_processo(driver)

        finally:
            output_file_name = os.path.splitext(os.path.basename(file_path))[0] + " - Resultado do Processamento.xlsx"
            output_file_path = os.path.join(processed_folder_path, output_file_name)
            df.to_excel(output_file_path, index=False)
            print(f"Resultados do processamento salvos em: {output_file_path}")

            try:
                shutil.move(file_path, processed_folder_path)
                print(f"Arquivo movido para a pasta: {processed_folder_path}")
            except Exception as e:
                print(f"Erro ao mover o arquivo para a pasta 'Processados': {e}")
                print(str(e))

            driver.quit()
