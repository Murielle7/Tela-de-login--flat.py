import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import pyautogui
import pyperclip
import pygetwindow as gw
from bs4 import BeautifulSoup
import win32gui
import win32con

def copiar_dados_planilha(origem, destino):
    try:
        df = pd.read_excel(origem, sheet_name='Planilha1', header=0)
        print(df.head())
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

    colunas_necessarias = ['Número Guia']
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print(f"Erro: uma ou mais colunas não foram encontradas na planilha 'Planilha1'.")
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
        
        print("Login realizado com sucesso.")
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
            EC.presence_of_element_located((By.XPATH, "/html/body/div[8]/div[3]/ul[3]/li/ul/li[1]/ul/li[2]/a"))
        )
        debitos_submenu.click()
        print("Clicou no submenu 'Financeiro - Consultar Guias'.")
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

def copiar_numero_guia(driver, numero_guia):
    try:
        driver.switch_to.default_content()

        # Acesse todos os iframes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        
        for iframe in iframes:
            driver.switch_to.frame(iframe)  # Tenta acessar o iframe
            try:
                guia_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='numeroGuiaCompleto']"))
                )
                guia_input.clear()
                guia_input.send_keys(numero_guia)
                print(f"Colou o número do processo: {numero_guia}")
                driver.switch_to.default_content()  # Retorna ao conteúdo principal
                return
            except Exception as e:
                print(f"Elemento não encontrado nesse iframe: {e}")
                driver.switch_to.default_content()  # Retorna ao conteúdo principal para tentar o próximo iframe

        print("Número da guia não encontrado dentro dos iframes ou na página atual.")
    except Exception as e:
        print(f"Erro ao colar o número da guia: {e}")

def clicar_botao_consultar(driver):
    try:
        driver.switch_to.default_content()  # Volta ao contexto principal

        # Acesse todos os iframes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        
        for iframe in iframes:
            driver.switch_to.frame(iframe)  # Tenta acessar o iframe
            try:
                botao_consultar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Consultar')]"))
                )
                botao_consultar.click()
                print("Botão 'Consultar' clicado.")
                driver.switch_to.default_content()  # Retorna ao conteúdo padrão
                return
            except Exception as e:
                print(f"Erro ao tentar clicar no botão dentro do iframe: {e}")
                driver.switch_to.default_content()  # Retorna ao conteúdo principal

    except Exception as e:
        print(f"Erro ao acessar o botão 'Consultar': {e}")

def clicar_botao_editar(driver):
    try:
        driver.switch_to.default_content()  # Volta ao contexto principal

        # Acesse todos os iframes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        
        for iframe in iframes:
            driver.switch_to.frame(iframe)  # Tenta acessar o iframe
            try:
                botao_editar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/table/tbody/tr/td[12]/button/i"))
                )
                botao_editar.click()
                print("Clicou em 'Editar Débito'.")
                driver.switch_to.default_content()  # Retorna ao conteúdo principal
                return
            except Exception as e:
                print(f"Erro ao tentar clicar no botão dentro do iframe: {e}")
                driver.switch_to.default_content()  # Retorna ao conteúdo principal

    except Exception as e:
        print(f"Erro ao acessar o botão 'Editar': {e}")

def focar_janela_navegador(title_part):
    try:
        windows = [win for win in gw.getAllTitles() if title_part in win]
        
        if windows:
            window = gw.getWindowsWithTitle(windows[0])[0]
            if window.isMinimized:
                window.restore()
            window.activate()
            print(f"Janela do navegador '{title_part}' em foco.")
        else:
            print(f"Janela com título contendo '{title_part}' não encontrada.")
    except Exception as e:
        print(f"Erro ao focar a janela do navegador: {e}")

def capturar_valores(driver):
    try:
        # Mude para o iframe específico
        driver.switch_to.frame(driver.find_element(By.ID, 'Principal'))

        # Obtenha o conteúdo do iframe
        iframe_content = driver.page_source
        soup = BeautifulSoup(iframe_content, 'html.parser')

        valores = {}

        # Extração de 'Valor Originário da Guia' (dentro do <b> do 2° <td> no 1° <tr>)
        valor_originario_elem = soup.select_one(
            'div#divCorpo.divCorpo > form#GuiaPreviaCalculo > div#divEditar.divEditar > '
            'fieldset.formEdicao > fieldset.VisualizaDados > table#Tabela.Tabela > tbody > '
            'tr:nth-of-type(1) > td:nth-of-type(2) > b'
        )
        valores['valor_original'] = valor_originario_elem.get_text(strip=True) if valor_originario_elem else "Não encontrado"
        print(f"'Valor Originário da Guia': {valores['valor_original']}")

        # SELETORES PARA 'JUROS FUNDESP-PJ S F(Reg.JURO)'
        seletores_juros = [
            ('div#divCorpo.divCorpo > form#GuiaPreviaCalculo > div#divEditar.divEditar > '
             'fieldset.formEdicao > fieldset.VisualizaDados > table#Tabela.Tabela > #tabListaEscala1> '
             'tr:nth-child(1) > td:nth-child(6)'),
            ('div#divCorpo.divCorpo > form#GuiaPreviaCalculo > div#divEditar.divEditar > '
             'fieldset.formEdicao > fieldset.VisualizaDados > table#Tabela.Tabela > #tabListaEscala1> '
             'tr:nth-child(1) > td:nth-child(5)')
        ]
        valores['juros'] = "Não encontrado"
        for seletor in seletores_juros:
            juros_fundesp_elem = soup.select_one(seletor)
            if juros_fundesp_elem:
                valores['juros'] = juros_fundesp_elem.get_text(strip=True)
                break
        print(f"'JUROS FUNDESP-PJ S F(Reg.JURO)': {valores['juros']}")
        
        # Extração de 'MULTA FUNDESP/PJ(Reg.M.FUNDESP)'
        seletores_multa = [
            ('div#divCorpo.divCorpo > form#GuiaPreviaCalculo > div#divEditar.divEditar > '
             'fieldset.formEdicao > fieldset.VisualizaDados > table#Tabela.Tabela > #tabListaEscala1> '
             'tr:nth-child(2) > td:nth-child(6)'),
            ('div#divCorpo.divCorpo > form#GuiaPreviaCalculo > div#divEditar.divEditar > '
             'fieldset.formEdicao > fieldset.VisualizaDados > table#Tabela.Tabela > #tabListaEscala1> '
             'tr:nth-child(2) > td:nth-child(5)')
        ]
        valores['multa'] = "Não encontrado"
        for seletor in seletores_multa:
            multa_fundesp_elem = soup.select_one(seletor)
            if multa_fundesp_elem:
                valores['multa'] = multa_fundesp_elem.get_text(strip=True)
                break
        print(f"'MULTA FUNDESP/PJ(Reg.M.FUNDESP)': {valores['multa']}")

        # Extração de 'Total Guia Atualizado (24/09/2024)'
        total_guia_atualizado_elem = soup.select_one(
            'div#divCorpo.divCorpo > form#GuiaPreviaCalculo > div#divEditar.divEditar > '
            'fieldset.formEdicao > fieldset.VisualizaDados > table#Tabela.Tabela > tbody:nth-child(6) > '
            'tr:nth-of-type(1) > td:nth-of-type(2) > b'
        )
        valores['total_guia'] = total_guia_atualizado_elem.get_text(strip=True) if total_guia_atualizado_elem else "Não encontrado"
        print(f"'Total Guia Atualizado (24/09/2024)': {valores['total_guia']}")

        return valores

    except Exception as e:
        print(f"Erro durante a extração: {e}")
        return None
        
def focar_janela_excel():
    try:
        excel_windows = [win for win in gw.getAllTitles() if 'Excel' in win]
        
        if excel_windows:
            excel_window = gw.getWindowsWithTitle(excel_windows[0])[0]
            if excel_window.isMinimized:
                excel_window.restore()
            excel_window.activate()
            print("Janela do Excel em foco.")
        else:
            print("Janela do Excel não encontrada.")
    except Exception as e:
        print(f"Erro ao focar a janela do Excel: {e}")

def colar_dados_na_planilha(linha, coluna, valor):
    try:
        focar_janela_excel() # Assegura que a janela do Excel está em foco
        time.sleep(0.5) # Pequena pausa para mudança de foco
        pyautogui.hotkey('ctrl', 'g')
        time.sleep(0.5)
        pyautogui.write(f"{coluna}{linha + 1}")
        pyautogui.press('enter')
        time.sleep(0.5)
        
        pyperclip.copy(valor)
        pyautogui.hotkey('ctrl', 'v')
        print(f"Valor '{valor}' colado na célula {coluna}{linha + 1}.")
        
    except Exception as e:
        print(f"Erro ao colar dados na planilha: {e}")

def salvar_planilha():
    try:
        focar_janela_excel()
        time.sleep(1)  # Pausa adequada antes de salvar
        pyautogui.hotkey('ctrl', 'b')  # Usa Ctrl + S para salvar
        time.sleep(1)
        print("Planilha salva.")
    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")

def clicar_botao_voltar(driver):
    try:
        driver.switch_to.default_content()  # Volta ao contexto principal

        # Acesse todos os iframes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        
        for iframe in iframes:
            driver.switch_to.frame(iframe)  # Tenta acessar o iframe
            try:
                botao_voltar = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div/div[1]/button[1]"))
                )
                botao_voltar.click()
                print("Botão 'Voltar' clicado.")
                driver.switch_to.default_content()  # Retorna ao conteúdo padrão
                return
            except Exception as e:
                print(f"Erro ao tentar clicar no botão dentro do iframe: {e}")
                driver.switch_to.default_content()  # Retorna ao conteúdo principal

    except Exception as e:
        print(f"Erro ao acessar o botão Voltar': {e}")

def main():
    # Caminhos do arquivo
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/VERDE -  P5.xlsx'
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
    
    clicar_analista_financeiro(driver)
    clicar_menu_cadastros(driver)

    # Iterar sobre os números de guia na tabela
    for index, row in df_filtrado.iterrows():
        numero_guia = row['Número Guia']

        try:
            copiar_numero_guia(driver, numero_guia)  # Copia o número da guia
            clicar_botao_consultar(driver)          # Clica no botão 'Consultar'
            clicar_botao_editar(driver)             # Clica no botão 'Editar'
            focar_janela_navegador("Projudi")
            focar_janela_excel()

            # Captura os valores necessários da página
            valores_extraidos = capturar_valores(driver)

            if valores_extraidos:
                colar_dados_na_planilha(index + 1, 'F', valores_extraidos.get('valor_original', ''))
                colar_dados_na_planilha(index + 1, 'G', valores_extraidos.get('juros', ''))
                colar_dados_na_planilha(index + 1, 'H', valores_extraidos.get('multa', ''))
                colar_dados_na_planilha(index + 1, 'I', valores_extraidos.get('total_guia', ''))

            salvar_planilha()  # Salva a planilha após cada iteração
            clicar_botao_voltar(driver)
            
            print("Reiniciando para o próximo número do processo...\n")
            time.sleep(1)

        except Exception as e:
            print(f"Erro no processamento do número guia {numero_guia}: {e}")

    driver.quit()  # Certifique-se de fechar o navegador ao final do processo

if __name__ == "__main__":
    main()
