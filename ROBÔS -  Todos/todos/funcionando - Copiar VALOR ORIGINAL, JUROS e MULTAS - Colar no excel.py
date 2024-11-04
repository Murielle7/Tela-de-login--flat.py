import time
from datetime import datetime  # Certifique-se de que esta linha está presente
import pygetwindow as gw
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
import pyperclip
import pyautogui
from bs4 import BeautifulSoup

def acessar_url(driver, url):
    driver.get(url)
    print("URL acessada.")

def copiar_nosso_numero(driver, nosso_numero):
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

        processo_input = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        processo_input.clear()
        processo_input.send_keys(nosso_numero)
        print(f"Colou o número do processo: {nosso_numero}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")
    finally:
        driver.switch_to.default_content()

def clicar_botao_buscar(driver):
    try:
        botao_buscar = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[5]/input[1]'))
        )
        botao_buscar.click()
        print("Botão 'Buscar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'buscar': {e}")

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

def buscar_numero_guia(driver, numero_guia):
    try:
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        
        pyautogui.hotkey('ctrl', 'a')   
        time.sleep(0.5)
        
        pyautogui.press('backspace')   
        time.sleep(0.5)
        
        pyautogui.write(str(numero_guia))
        pyautogui.press('enter')
        print(f"Número da guia '{numero_guia}' digitado no navegador.")
    except Exception as e:
        print(f"Erro ao digitar número guia: {e}")

def localizar_e_copiar_valores_por_texto(driver):
    try:
        # Obtenha o texto completo da página
        page_source = driver.page_source
        print("Página capturada")

        # Analise o HTML com BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')
        
        valores = {}

        # Valor Original
        valor_original = soup.find(text="Valor Originário da Guia")
        if valor_original:
            valor_original_td = valor_original.find_next('td', align="right")
            if valor_original_td:
                valores["valor_original"] = valor_original_td.text.strip()

        # JUROS FUNDESP-PJ S F
                juros_row = soup.find('td', text="JUROS FUNDESP-PJ S F(Reg.JURO)")
        if juros_row:
            juros_td = juros_row.find_next('td', align="right")
            if juros_td:
                valores["juros"] = juros_td.text.strip()

        # MULTA FUNDESP/PJ
        multa_row = soup.find('td', text="MULTA FUNDESP/PJ(Reg.M.FUNDESP)")
        if multa_row:
            multa_td = multa_row.find_next('td', align="right")
            if multa_td:
                valores["multa"] = multa_td.text.strip()

        # Total Guia Atualizado
        total_guia_text = soup.find(text=re.compile("Total Guia Atualizado"))
        if total_guia_text:
            total_guia_td = total_guia_text.find_next('td', align="right")
            if total_guia_td:
                valores["total_guia"] = total_guia_td.text.strip()

        for key, value in valores.items():
            print(f"{key} encontrado: {value}")
        
        return valores

    except Exception as e:
        print(f"Erro ao copiar os valores: {e}")
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
        pyautogui.hotkey('ctrl', 'b')
        time.sleep(1)
        print("Planilha salva.")
    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")

def main():
    # Aguarda até às 17:35 de 13/09/2024
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/Demanda DANILO - ALTERAR VALORES (JUNÇÃO).xlsx'
    nome_planilha = 'verde'
    tabela_boleto = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, skiprows=0)

    url = 'https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=4&TipoConsultaProcesso=24&ServletRedirect=GuiaFinal_FinalZeroPublica&TituloDaPagina=Guia+Final+e+Final+Zero+%5BAcesso+P%C3%BAblico%5D&hashFluxo=1722628305255'
    
    options = webdriver.FirefoxOptions()
    with webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options) as driver:
        acessar_url(driver, url)

        for index, row in tabela_boleto.iterrows():
            nosso_numero = row['Número Processo']
            numero_guia = row['Número Guia']
            
            copiar_nosso_numero(driver, nosso_numero)
            clicar_botao_buscar(driver)

            focar_janela_navegador("Projudi")
            buscar_numero_guia(driver, numero_guia)

            valores = localizar_e_copiar_valores_por_texto(driver)
            print(f"Valores extraídos: {valores}")
            if valores:
                colar_dados_na_planilha(index + 1, 'F', valores.get('valor_original', ''))
                colar_dados_na_planilha(index + 1, 'G', valores.get('juros', ''))
                colar_dados_na_planilha(index + 1, 'H', valores.get('multa', ''))
                colar_dados_na_planilha(index + 1, 'I', valores.get('total_guia', ''))

            salvar_planilha()

            print("Reiniciando para o próximo número do processo...\n")
            time.sleep(1)

            acessar_url(driver, url)

if __name__ == "__main__":
    main()
