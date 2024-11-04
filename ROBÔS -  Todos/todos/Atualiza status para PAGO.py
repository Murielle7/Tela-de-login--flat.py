import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager

def copiar_dados_planilha(origem, destino):
    try:
        df = pd.read_excel(origem, sheet_name='Conferência - GO', header=3)
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None
    
    colunas_necessarias = ['Número Processo', 'Número Guia', 'Status', 'ID']
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print(f"Erro: uma ou mais colunas não foram encontradas na planilha 'Conferência - GO'.")
        return None
    
    df_filtrado = df[colunas_necessarias]
    
    try:
        df_filtrado.to_excel(destino, index=False)
        print(f"Dados salvos com sucesso em: {destino}")
    except Exception as e:
        print(f"Erro ao salvar a nova planilha: {e}")
    
    return df_filtrado

def abrir_site_e_login(url, username, password):
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
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
        analista_link = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Analista Financeiro')]"))
        )
        analista_link.click()
        print("Clicou em 'Analista Financeiro'.")
    except Exception as e:
        print(f"Erro ao clicar em 'Analista Financeiro': {e}")

def clicar_menu_cadastros(driver):
    try:
        cadastros_menu = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "2"))
        )
        cadastros_menu.click()
        print("Clicou no menu 'Cadastros'.")
        
        financeiros_submenu = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))
        )
        financeiros_submenu.click()
        print("Clicou no submenu 'Financeiros'.")

        debitos_submenu = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))
        )
        debitos_submenu.click()
        print("Clicou no submenu 'Débitos - Diretoria Financeira'.")
        
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

def teclar_seta_para_cima(driver, vezes=2):
    try:
        actions = webdriver.ActionChains(driver)
        for _ in range(vezes):
            actions.send_keys(Keys.ARROW_UP)
        actions.perform()
        print(f"Teclou Seta para Cima {vezes} vezes.")
    except Exception as e:
        print(f"Erro ao teclar Seta para Cima: {e}")

def copiar_numero_processo(driver, numero_processo):
    try:
        # Verificar se o campo está dentro de um iframe
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

        # Aguardar até que o campo do número do processo esteja presente e clicável
        processo_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        time.sleep(1)  # Esperar um tempo adicional para garantir que o campo esteja interativo
        processo_input.clear()  # Limpar o campo antes de colar o novo número
        processo_input.send_keys(numero_processo)  # Inserir o número do processo
        print(f"Colou o número do processo: {numero_processo}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

def clicar_botao_consultar(driver):
    try:
        # Localize o botão "Consultar" usando o FullPath e clique nele
        botao_consultar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/div/input"))
        )
        botao_consultar.click()
        print("Botão 'Consultar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Consultar': {e}")

def localizar_id_e_clicar_editar(driver, id_processo):
    try:
        # Aguardar até que a tabela esteja presente
        tabela = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        
        # Procurar a linha que contém o ID do processo
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                id_text = coluna.text.strip().replace(".", "")
                if id_text == str(int(id_processo)):  # Converter ID para inteiro e depois para string
                    print(f"ID {id_processo} encontrado na tabela.")
                    # Tentar localizar o botão "Editar Débito" dentro da mesma linha
                    botao_editar = linha.find_element(By.XPATH, ".//input[@type='image' and contains(@title, 'Editar Débito')]")
                    botao_editar.click()
                    print(f"Clicou em 'Editar Débito' para o ID: {id_processo}")
                    return
        print(f"ID do processo {id_processo} não encontrado na tabela.")
    except Exception as e:
        print(f"Erro ao localizar o ID do processo e clicar em 'Editar Débito': {e}")

def selecionar_status(driver, status_processo):
    try:
        # Aguardar até que a caixa de seleção esteja presente
        select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "Id_ProcessoDebitoStatus"))
        )
        
        # Criar um objeto Select a partir do elemento
        select = Select(select_element)
        
        # Selecionar a opção pelo texto visível
        select.select_by_visible_text(status_processo)
        print(f"Selecionou o status: {status_processo}")
    except Exception as e:
        print(f"Erro ao selecionar o status: {e}")

def clicar_botao_salvar(driver):
    try:
        botao_salvar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divPortaBotoes"]/button[1]'))
        )
        botao_salvar.click()
        print("Botão 'Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Salvar': {e}")

def clicar_botao_confirmar_salvar(driver):
    try:
        botao_confirmar_salvar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divConfirmarSalvar"]/button'))
        )
        botao_confirmar_salvar.click()
        print("Botão 'Confirmar Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Confirmar Salvar': {e}")

def clicar_botao_ok(driver):
    try:
        # Adicionar espera explícita para garantir que a ação anterior seja concluída
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

def voltar_para_pagina_pesquisa(driver):
    try:
        driver.back()  # Voltar para a página anterior
        print("Voltar uma página.")
    except Exception as e:
        print(f"Erro ao clicar no botão'Voltar uma página': {e}")

def main():
    caminho_arquivo_origem = r'\\10.0.20.65\Financeira\CAJ\Cobrança - Judicial e Administrativa\Custas Judiciais - Finais\Remessas - CRA\Remessa.xlsx'
    caminho_arquivo_destino = os.path.join(os.path.expanduser('~'), 'Desktop', 'Teste Python 123.xlsx')
    
    df = copiar_dados_planilha(caminho_arquivo_origem, caminho_arquivo_destino)
    if df is None:
        return
    
    url = 'https://projudi.tjgo.jus.br/'
    username = '02131786131'  # Substitua pelo seu usuário
    password = 'Gabriel*8'  # Substitua pela sua senha
    
    driver = abrir_site_e_login(url, username, password)
    
    if driver is None:
        return
    
    clicar_analista_financeiro(driver)
    clicar_menu_cadastros(driver)
    teclar_seta_para_cima(driver, 2)
    
    for index, row in df.iterrows():
        numero_processo = row['Número Processo']
        status_processo = row['Status']
        id_processo = row['ID']
        
        copiar_numero_processo(driver, numero_processo)
        clicar_botao_consultar(driver)
        
        localizar_id_e_clicar_editar(driver, id_processo)
        
        # Selecionar o status do processo
        selecionar_status(driver, status_processo)
        
        # Clicar no botão "Salvar"
        clicar_botao_salvar(driver)
        
        # Clicar no botão "Confirmar Salvar"
        clicar_botao_confirmar_salvar(driver)
        
        # Clicar no botão "Ok"
        clicar_botao_ok(driver)
        
        time.sleep(2)  # Adicionar um atraso para garantir que todas as ações sejam realizadas
        
    input("Pressione ENTER para fechar o navegador e finalizar o script...")
    driver.quit()

if __name__ == "__main__":
    main()
