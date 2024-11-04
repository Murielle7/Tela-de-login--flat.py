import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager

def copiar_dados_planilha(origem, destino):
    try:
        df = pd.read_excel(origem, sheet_name='Export', header=0)
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None
    
    colunas_necessarias = ['Número Processo', 'Número Guia']
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print(f"Erro: uma ou mais colunas não foram encontradas na planilha 'Export'.")
        return None

    df_filtrado = df[colunas_necessarias].copy()
    df_filtrado['Alteração de Status'] = ""  # Adiciona a nova coluna para status

    try:
        df_filtrado.to_excel(destino, index=False)
        print(f"Dados salvos com sucesso em: {destino}")
    except Exception as e:
        print(f"Erro ao salvar a nova planilha: {e}")

    return df_filtrado

def abrir_site_e_login(url, username, password):
    try:
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service)
        driver.get(url)

        user_field = WebDriverWait(driver, 3).until(
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

def clicar_menu_cadastros(driver):
    try:
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "2"))  # ID do menu 'Cadastros'
        ).click()
        print("Clicou no menu 'Cadastros'.")
        
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))  # Link de texto 'Financeiros'
        ).click()
        print("Clicou no submenu 'Financeiros'.")

        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))  # Link de texto 'Débitos - Diretoria Financeira'
        ).click()
        print("Clicou no submenu 'Débitos - Diretoria Financeira'.")
        
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

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

        processo_input = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        time.sleep(2)
        processo_input.clear()
        processo_input.send_keys(numero_processo)
        print(f"Colou o número do processo: {numero_processo}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

def clicar_botao_consultar(driver):
    try:
        botao_consultar = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='Consultar']"))
        )
        botao_consultar.click()
        print("Botão 'Consultar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Consultar': {e}")

def localizar_guias_e_xpaths(driver, numero_guia):
    try:
        guias_info = []
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset[2]/table"))  # Certificar existência da tabela
        )
        
        page_html = driver.page_source
        soup = BeautifulSoup(page_html, 'html.parser')
        
        linhas = soup.select("table#Tabela tbody tr")
        for linha_idx, linha in enumerate(linhas):
            colunas = linha.find_all("td")
            if len(colunas) < 3:
                continue
            
            guia_text = colunas[2].get_text(strip=True).replace(".", "").replace("/", "").replace("-", "")
            
            if guia_text == numero_guia.replace(".", "").replace("/", "").replace("-", ""):
                editar_botao_xpath = f"(/html/body/div[1]/form/div[2]/fieldset[2]/table/tbody/tr[{linha_idx + 1}]/td[8]/a)[1]"
                status_text = colunas[1].get_text(strip=True)
                guias_info.append((guia_text, status_text, editar_botao_xpath))
                
        return guias_info
    
    except Exception as e:
        print(f"Erro ao localizar o número da guia: {e}")
        return []

def extrair_xpath_status(driver):
    try:
        page_html = driver.page_source
        soup = BeautifulSoup(page_html, 'html.parser')

        options = soup.select('#Id_ProcessoDebitoStatus option')
        
        status_xpaths = {}
        for idx, option in enumerate(options):
            status_text = option.get_text(strip=True)
            status_xpath = f'//*[@id="Id_ProcessoDebitoStatus"]/option[{idx + 1}]'
            status_xpaths[status_text] = status_xpath
        
        return status_xpaths
    
    except Exception as e:
        print(f"Erro ao extrair os XPaths dos status: {e}")
        return {}

def alterar_status(driver, novo_status_xpath):
    try:
        status_dropdown = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Id_ProcessoDebitoStatus"]'))
        )
        status_dropdown.click()
        time.sleep(2)
        novo_status_opcao = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, novo_status_xpath))
        )
        novo_status_opcao.click()
        print(f"Status alterado para: {novo_status_opcao.text}")
    except Exception as e:
        print(f"Erro ao alterar o status: {e}")

def clicar_botao_salvar(driver):
    try:
        botao_salvar = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divPortaBotoes"]/button[1]'))
        )
        botao_salvar.click()
        print("Botão 'Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Salvar': {e}")

def clicar_botao_confirmar_salvar(driver):
    try:
        botao_confirmar_salvar = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divConfirmarSalvar"]/button'))
        )
        botao_confirmar_salvar.click()
        print("Botão 'Confirmar Salvar' clicado.")
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar no botão 'Confirmar Salvar': {e}")

def clicar_botao_ok(driver):
    try:
        botao_ok = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[3]/div/button'))
        )
        botao_ok.click()
        print("Botão 'Ok' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Ok': {e}")

def clicar_botao_escolher_novo_processo(driver):
    try:
        botao_novo_processo = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[1]/button[2]'))
        )
        botao_novo_processo.click()
        print("Botão 'Escolher Novo Processo' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Escolher Novo Processo': {e}")

def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/alterar status -  apto --- protestado.xlsx'
    caminho_arquivo_destino = os.path.join(os.path.expanduser('~'), 'Desktop', 'Teste Python 123.xlsx')
    
    df_filtrado = copiar_dados_planilha(caminho_arquivo, caminho_arquivo_destino)
    if df_filtrado is None:
        return
    
    url = 'https://projudi.tjgo.jus.br/'
    username = '70273080105'
    password = '235791Mu@'
    
    driver = abrir_site_e_login(url, username, password)
    
    if driver is None:
        return
    
    clicar_menu_cadastros(driver)
    
    for index, row in df_filtrado.iterrows():
        numero_processo = row['Número Processo']
        numero_guia = row['Número Guia']
        
        if pd.isna(numero_guia):
            print("Número da Guia vazio encontrado. Finalizando o processamento.")
            break

        try:
            copiar_numero_processo(driver, numero_processo)
            clicar_botao_consultar(driver)
            
            guias_info = localizar_guias_e_xpaths(driver, numero_guia)
            if not guias_info:
                print(f"Número da Guia {numero_guia} não encontrado para o processo {numero_processo}.")
                df_filtrado.loc[index, 'Alteração de Status'] = "Guia não encontrada"
                continue
            
            for guia_text, status_text, editar_xpath in guias_info:
                if status_text == "Em Análise pela Financeira":
                    driver.find_element(By.XPATH, editar_xpath).click()
                    time.sleep(2)
                    
                    status_xpaths = extrair_xpath_status(driver)
                    if "Protestado" in status_xpaths:
                        alterar_status(driver, status_xpaths["Protestado"])
                        clicar_botao_salvar(driver)
                        clicar_botao_confirmar_salvar(driver)
    
                    clicar_botao_ok(driver)
                    df_filtrado.loc[index, 'Alteração de Status'] = "Alterado com sucesso"
            
            clicar_botao_escolher_novo_processo(driver)
                
        except Exception as e:
            print(f"Erro ao processar o processo {numero_processo}: {e}")
            df_filtrado.loc[index, 'Alteração de Status'] = "Erro na Alteração"

    try:
        df_filtrado.to_excel(caminho_arquivo_destino, index=False)
        print(f"Planilha atualizada com os resultados salvos em: {caminho_arquivo_destino}")
    except Exception as e:
        print(f"Erro ao salvar a planilha atualizada: {e}")
    
    input("Pressione ENTER para fechar o navegador e finalizar o script...")
    driver.quit()

if __name__ == "__main__":
    main()
