import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.firefox import GeckoDriverManager

def copiar_dados_planilha(origem, destino):
    try:
        df = pd.read_excel(origem, sheet_name='Export', header=0)
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None
    
    colunas_necessarias = ['Número Processo', 'Número Guia', 'Status Débito']
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

def clicar_analista_financeiro(driver):
    try:
        analista_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Analista Financeiro')]"))
        )
        analista_link.click()
        print("Clicou em 'Analista Financeiro'.")
    except Exception as e:
        print(f"Erro ao clicar em 'Analista Financeiro': {e}")

def clicar_menu_cadastros(driver):
    try:
        cadastros_menu = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "2"))
        )
        cadastros_menu.click()
        print("Clicou no menu 'Cadastros'.")
        
        financeiros_submenu = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))
        )
        financeiros_submenu.click()
        print("Clicou no submenu 'Financeiros'.")

        debitos_submenu = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))
        )
        debitos_submenu.click()
        print("Clicou no submenu 'Débitos - Diretoria Financeira'.")
        
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

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
        processo_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        time.sleep(2)  # Esperar um tempo adicional para garantir que o campo esteja interativo
        processo_input.clear()  # Limpar o campo antes de colar o novo número
        processo_input.send_keys(numero_processo)  # Inserir o número do processo
        print(f"Colou o número do processo: {numero_processo}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

def clicar_botao_consultar(driver):
    try:
        # Localize o botão "Consultar" usando o FullPath e clique nele
        botao_consultar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/div/input"))
        )
        botao_consultar.click()
        print("Botão 'Consultar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Consultar': {e}")

def localizar_e_alterar_status_do_numero_guia(driver, numero_guia):
    try:
        # Limpar formatação do número da guia
        numero_guia = str(numero_guia).replace(".", "").replace("/", "").replace("-", "")
        
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        
        linhas = tabela.find_elements(By.TAG_NAME, "tr")  # Obtém todas as linhas da tabela
        total_guias_alteradas = 0  # Contador de guias alteradas

        # Percorre todas as linhas da tabela
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")  # Obtém todas as colunas da linha
            linha_texto = linha.text.strip()  # Captura o texto da linha
            print(f"Analisando linha: {linha_texto}")

            # Verifique se o texto contém o número da guia
            if numero_guia in linha_texto:
                print(f"NÚMERO DA GUIA {numero_guia} encontrado na linha.")

                # Tenta localizar todos os ícones de edição na linha
                editar_icons = linha.find_elements(By.XPATH, ".//input[@name='formLocalizarimgEditar' and @src='./imagens/imgEditarPequena.png']")
                
                # Clica em cada ícone de edição encontrado
                for editar_icon in editar_icons:
                    if editar_icon.is_displayed():  # Verifica se o ícone está visível
                        editar_icon.click()  # Clica para editar o débito
                        print("Ícone de edição clicado.")

                        # Altera o status para 'PROTESTADO'
                        selecionar_status_para_protestado(driver)  # Altera o status
                        clicar_botao_salvar(driver)  # Salva a alteração
                        clicar_botao_confirmar_salvar(driver)  # Confirma a alteração
                        clicar_botao_ok(driver)  # Clica 'Ok' para fechar a janela

                        total_guias_alteradas += 1  # Incrementa o contador de guias alteradas
                        print(f"Status alterado para 'PROTESTADO' para a guia: {numero_guia}")

                        # Retorna ao contexto original após cada alteração
                        driver.switch_to.default_content()
                        time.sleep(1)  # Espera para garantir que todas as interações foram completadas

                break  # Para sair após lidar com todas as edições dessa guia

        print(f"Número total de guias alteradas para 'PROTESTADO': {total_guias_alteradas}")

    except Exception as e:
        print(f"Erro ao localizar campos relacionados à NÚMERO DA GUIA: {e}")

def selecionar_status_para_protestado(driver):
    try:
        # Espera que o campo de seleção de status esteja presente e interage com ele
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/form/div[2]/fieldset[1]/select'))
        ).click()  # Clica no dropdown para abrir as opções

        opcao_protestado = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[2]/fieldset[1]/select/option[19]'))
        )
        opcao_protestado.click()  # Clica na opção "PROTESTADO"
        print("Selecionou a opção 'PROTESTADO'.")
        
    except Exception as e:
        print(f"Erro ao selecionar o status: {e}")

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
        
        # Consultar o processo antes de alterar o status
        copiar_numero_processo(driver, numero_processo)
        clicar_botao_consultar(driver)  # Clica no botão de consultar

        # Chama a função para alterar o status da guia
        localizar_e_alterar_status_do_numero_guia(driver, numero_guia)

def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/TESTE-PEDRO.xlsx'
    caminho_arquivo_destino = os.path.join(os.path.expanduser('~'), 'Desktop', 'Teste Python 123.xlsx')
    
    # Carregar dados da planilha
    df_filtrado = copiar_dados_planilha(caminho_arquivo, caminho_arquivo_destino)
    if df_filtrado is None:
        return  # Sai se a planilha não foi carregada
    
    # Informações de login e URL do Projudi
    url = 'https://projudi.tjgo.jus.br/'
    username = '70273080105'
    password = '235791Mu@'
    
    # Iniciar session do Selenium
    driver = abrir_site_e_login(url, username, password)
    if driver is None:
        return  # Sai se o driver não foi iniciado corretamente
    
    clicar_analista_financeiro(driver)
    clicar_menu_cadastros(driver)
    
    # Processar todas as ocorrências
    processar_todas_as_ocorrencias(driver, df_filtrado)
    
    input("Pressione ENTER para fechar o navegador e finalizar o script...")
    driver.quit()

if __name__ == "__main__":
    main()
