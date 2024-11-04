import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager

def ler_planilha(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name='Export', header=0)
        df_filtrado = df[(df['Status Débito'] == 'Apto para envio ao protesto') & 
                         (df['Ocorrência'].isin(['PROTESTADO', 'PROTESTO POR EDITAL']))]
        df_filtrado = df_filtrado[['Número Processo', 'Nome Parte']]
        return df_filtrado
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return None

def abrir_site_e_login(url, username, password):
    try:
        options = Options()
        options.headless = False
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service, options=options)
        driver.get(url)
        
        user_field = WebDriverWait(driver, 10).until(
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
        cadastros_menu = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "2"))
        )
        cadastros_menu.click()
        print("Clicou no menu 'Cadastros'.")
        
        financeiros_submenu = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))
        )
        financeiros_submenu.click()
        print("Clicou no submenu 'Financeiros'.")

        debitos_submenu = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))
        )
        debitos_submenu.click()
        print("Clicou no submenu 'Débitos - Diretoria Financeira'.")
        
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

def procedimento_para_cada_processo(driver, df):
    for index, row in df.iterrows():
        try:
            numero_processo = row['Número Processo']
            nome_parte = row['Nome Parte']
            
            copiar_numero_processo(driver, numero_processo)
            clicar_elemento(driver, By.XPATH, '/html/body/div[1]/form/div[2]/fieldset/div/input', "'Consultar'")
            
            if localizar_nome_parte_e_clicar_editar(driver, nome_parte):
                selecionar_status(driver)
                clicar_botao(driver, '//*[@id="divPortaBotoes"]/button[1]', 'Salvar')
                clicar_botao(driver, '//*[@id="divConfirmarSalvar"]/button', 'Confirmar Salvar')
                clicar_botao(driver, '/html/body/div[8]/div[3]/div/button', 'Ok')

            clicar_botao(driver, '/html/body/div[1]/form/div[1]/button[2]', 'Escolher Novo Processo')
        
        except Exception as e:
            print(f"Erro ao processar o processo {numero_processo}: {e}")

def copiar_numero_processo(driver, numero_processo):
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for iframe in iframes:
            driver.switch_to.frame(iframe)
            try:
                processo_input = driver.find_element(By.ID, "ProcessoNumero")
                if processo_input:
                    processo_input.clear()
                    processo_input.send_keys(numero_processo)
                    print(f"Colou o número do processo: {numero_processo}")
                    driver.switch_to.default_content()  # Return to main content after interacting with iframe
                    return
            except Exception:
                driver.switch_to.default_content()
                continue
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

def localizar_nome_parte_e_clicar_editar(driver, nome_parte):
    try:
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                if coluna.text.strip() == nome_parte:
                    print(f"Nome {nome_parte} encontrado na tabela.")
                    botao_editar = linha.find_element(By.XPATH, ".//input[@type='image' and contains(@title, 'Editar Débito')]")
                    botao_editar.click()
                    print(f"Clicou em 'Editar Débito' para Nome Parte: {nome_parte}")
                    return True
        print(f"Nome Parte {nome_parte} não encontrado na tabela.")
        return False
    except Exception as e:
        print(f"Erro ao localizar Nome Parte e clicar em 'Editar Débito': {e}")
        return False

def selecionar_status(driver):
    try:
        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Id_ProcessoDebitoStatus"))
        )
        opcao_protestado = select_element.find_element(By.XPATH, '//option[text()="PROTESTADO"]')
        opcao_protestado.click()
        print("Selecionou a opção 'PROTESTADO'")
    except Exception as e:
        print(f"Erro ao selecionar o status do processo: {e}")

def clicar_botao(driver, xpath, descricao_botao):
    try:
        botao = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        botao.click()
        print(f"Botão '{descricao_botao}' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão '{descricao_botao}': {e}")

def main():
    caminho_arquivo = r'S:/CAJ/Cobrança - Judicial e Administrativa/Custas Judiciais - Finais/Processamento Retorno - CRA/Aguardando Processamento/Relatório BI - Copia.xlsx'
    df_filtrado = ler_planilha(caminho_arquivo)
    if df_filtrado is None:
        return
    
    url = 'https://projudi.tjgo.jus.br/'
    username = '70273080105'  # replace with actual username
    password = '235791Mu@'   # replace with actual password
    
    driver = abrir_site_e_login(url, username, password)
    if driver is None:
        return
    
    clicar_menu_cadastros(driver)
    procedimento_para_cada_processo(driver, df_filtrado)
    localizar_nome_parte_e_clicar_editar(driver, nome_parte)
    selecionar_status(driver)
    clicar_botao_confirmar_salvar(driver)
    clicar_botao_ok(driver)
    clicar_botao_escolher_novo_processo(driver)
    
    input("Pressione ENTER para fechar o navegador e finalizar o script...")
    driver.quit()

if __name__ == "__main__":
    main()
