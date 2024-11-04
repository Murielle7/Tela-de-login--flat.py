import os
import time
import pandas as pd
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
    
    colunas_necessarias = ['Número Processo', 'Número Guia', 'Status Débito']
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print("Erro: uma ou mais colunas não foram encontradas na planilha 'Export'.")
        return None
    
    df_filtrado = df[colunas_necessarias].copy()
    df_filtrado['Alteração de Status'] = ""

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

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "login"))
        ).send_keys(username)
        
        driver.find_element(By.ID, "senha").send_keys(password)
        driver.find_element(By.ID, "senha").send_keys(Keys.RETURN)
        
        return driver
    except Exception as e:
        print(f"Erro ao abrir o site e realizar login: {e}")
        return None

def clicar_analista_financeiro(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Analista Financeiro')]"))
        ).click()
        print("Clicou em 'Analista Financeiro'.")
    except Exception as e:
        print(f"Erro ao clicar em 'Analista Financeiro': {e}")

def clicar_menu_cadastros(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "2"))
        ).click()
        print("Clicou no menu 'Cadastros'.")
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))
        ).click()
        print("Clicou no submenu 'Financeiros'.")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))
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
                driver.find_element(By.ID, "ProcessoNumero").clear()
                driver.find_element(By.ID, "ProcessoNumero").send_keys(numero_processo)
                print(f"Colou o número do processo: {numero_processo}")
                driver.switch_to.default_content()
                break
            except:
                driver.switch_to.default_content()
                continue
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

def clicar_botao_consultar(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/div/input"))
        ).click()
        print("Botão 'Consultar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Consultar': {e}")

def localizar_guia_e_clicar_editar_todos(driver, numero_guia):
    try:
        numero_guia = str(numero_guia).replace(".", "").replace("/", "").replace("-", "")
                tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        guias_encontradas = 0

        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                guia_text = str(coluna.text).strip().replace(".", "").replace("/", "").replace("-", "")
                if guia_text == numero_guia:
                    print(f"NÚMERO DA GUIA {numero_guia} encontrado na tabela.")
                    botao_editar = linha.find_element(By.XPATH, ".//input[@type='image' and contains(@title, 'Editar Débito')]")
                    botao_editar.click()

                    # Realiza as ações de seleção de status e salvamento
                    selecionar_status(driver)
                    clicar_botao_salvar(driver)
                    clicar_botao_confirmar_salvar(driver)
                    clicar_botao_ok(driver)
                    
                    guias_encontradas += 1

                    # Voltar para o estado padrão para continuar a operar na tabela
                    driver.switch_to.default_content()
                    tabela = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
                    )
                    linhas = tabela.find_elements(By.TAG_NAME, "tr")
                    break  # Continue verificando a partir do próximo guia

        if guias_encontradas == 0:
            print(f"NÚMERO DA GUIA {numero_guia} não encontrado na tabela.")
            return False
        else:
            return True
    except Exception as e:
        print(f"Erro ao localizar o NÚMERO DA GUIA e clicar em 'Editar Débito': {e}")
        return False

def selecionar_status(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Id_ProcessoDebitoStatus"))
        )

        opcao_em_analise = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Id_ProcessoDebitoStatus"]/option[text()="PROTESTADO"]'))
        )
        opcao_em_analise.click()
        print("Selecionou a opção 'PROTESTADO'")
    except Exception as e:
        print(f"Erro ao inserir os dados no campo e realizar as ações: {e}")

def clicar_botao_salvar(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divPortaBotoes"]/button[1]'))
        ).click()
        print("Botão 'Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Salvar': {e}")

def clicar_botao_confirmar_salvar(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divConfirmarSalvar"]/button'))
        ).click()
        print("Botão 'Confirmar Salvar' clicado.")
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar no botão 'Confirmar Salvar': {e}")

def clicar_botao_ok(driver):
    try:
        botao_ok = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[3]/div/button'))
        )
        botao_ok.click()
        print("Botão 'Ok' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Ok': {e}")

def processar_todas_as_ocorrencias(driver, df_filtrado):
    for index, row in df_filtrado.iterrows():
        numero_processo = row['Número Processo']
        status_processo = row['Status Débito']
        numero_guia = row['Número Guia']
        
        if pd.isna(numero_guia):
            print("Número GUIA vazio encontrado.")
            continue
        
        numero_guia_encontrado = localizar_guia_e_clicar_editar_todos(driver, numero_guia)

        if numero_guia_encontrado:
            df_filtrado.loc[index, 'Alteração de Status'] = "Alterado com sucesso"
            print(f"Status alterado para todas as ocorrências do NÚMERO DA GUIA: {numero_guia}")
        else:
            df_filtrado.loc[index, 'Alteração de Status'] = "Erro na Alteração"
            print(f"Erro na alteração para o NÚMERO DA GUIA: {numero_guia}")

            caminho_arquivo_destino = os.path.join(os.path.expanduser('~'), 'Desktop', 'Teste Python 123.xlsx')

    try:
        df_filtrado.to_excel(caminho_arquivo_destino, index=False)
        print(f"Planilha atualizada com os resultados salvos em: {caminho_arquivo_destino}")
    except Exception as e:
        print(f"Erro ao salvar a planilha atualizada: {e}")

def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/data (14).xlsx'
    caminho_arquivo_destino = os.path.join(os.path.expanduser('~'), 'Desktop', 'Teste Python 123.xlsx')
    
    # Carregar dados da planilha
    df_filtrado = copiar_dados_planilha(caminho_arquivo, caminho_arquivo_destino)
    if df_filtrado is None:
        return
    
    # Informações de login e URL do Projudi
    url = 'https://projudi.tjgo.jus.br/'
    username = '70273080105'
    password = '235791Mu@'
    
    # Iniciar session do Selenium
    driver = abrir_site_e_login(url, username, password)
    if driver is None:
        return
    
    clicar_analista_financeiro(driver)
    clicar_menu_cadastros(driver)
    
    # Processar todas as ocorrências
    processar_todas_as_ocorrencias(driver, df_filtrado)
    
    input("Pressione ENTER para fechar o navegador e finalizar o script...")
    driver.quit()

if __name__ == "__main__":
    main()



        
