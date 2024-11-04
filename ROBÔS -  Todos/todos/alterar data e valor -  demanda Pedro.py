import pyautogui
import pyperclip
import pygetwindow as gw
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Função para copiar o número do processo em um campo
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

        processo_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        time.sleep(1)
        processo_input.clear()
        processo_input.send_keys(nosso_numero)
        print(f"Colou o número do processo: {nosso_numero}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")
    finally:
        driver.switch_to.default_content()

# Função para clicar no botão "Buscar"
def clicar_botao_buscar(driver):
    try:
        botao_buscar = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[5]/input[1]'))
        )
        botao_buscar.click()
        print("Botão 'Buscar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'buscar': {e}")

# Função para clicar nos botões "Emitir e Imprimir"
def clicar_emitir_imprimir(driver):
    try:
        botao_emitir_boleto = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="imgEmitirGuiaImprimirPDF"]'))
        )
        botao_emitir_boleto.click()
        time.sleep(1)

        botao_emitir_final = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[9]/fieldset/div/button'))
        )
        botao_emitir_final.click()
        time.sleep(1)  # Aguarda mais tempo para garantir que a página seja carregada
    except Exception as e:
        print(f"Erro ao clicar em Emitir e Imprimir: {e}")

# Função para extrair dados da guia online usando pyautogui
def extrair_dados_guia():
    try:
        # Extrair data de vencimento ORIGINAL
        x_start, y_start = 184, 378  # Ponto inicial da seleção
        x_end, y_end = 262, 378      # Ponto final da seleção
        
        # Move o cursor para a posição inicial e faz a seleção
        pyautogui.moveTo(x_start, y_start)
        pyautogui.mouseDown()
        pyautogui.moveTo(x_end, y_end)
        pyautogui.mouseUp()

        # Copia o texto selecionado
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(1)  # Aguarda um pouco mais para garantir que o texto foi copiado

        # Obtém o texto da área copiada
        data_vencimento = pyperclip.paste().strip()

        # Extrair valor atualizado
        x_start, y_start = 829, 531  # Ponto inicial da seleção
        x_end, y_end = 871, 531      # Ponto final da seleção
        
        # Move o cursor para a posição inicial e faz a seleção
        pyautogui.moveTo(x_start, y_start)
        pyautogui.mouseDown()
        pyautogui.moveTo(x_end, y_end)
        pyautogui.mouseUp()

        # Copia o texto selecionado
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(1)  # Aguarda um pouco mais para garantir que o texto foi copiado

        # Obtém o texto da área copiada
        valor_atualizado = pyperclip.paste().strip()

        print(f"Data de vencimento extraída: {data_vencimento}")
        print(f"Valor atualizado extraído: {valor_atualizado}")
        
        return data_vencimento, valor_atualizado
    except Exception as e:
        print(f"Erro ao extrair dados da guia: {e}")
        return None, None

# Função para focar a janela do Excel
def focar_janela_excel():
    try:
        # Encontra todas as janelas cujo título contém "Excel"
        excel_windows = [win for win in gw.getAllTitles() if 'Excel' in win]
        
        if excel_windows:
            # Seleciona a primeira janela encontrada
            excel_window = gw.getWindowsWithTitle(excel_windows[0])[0]
            if not excel_window.isMinimized:
                excel_window.minimize()
            excel_window.restore()
            excel_window.activate()
            print("Janela do Excel em foco.")
        else:
            print("Janela do Excel não encontrada.")
    except Exception as e:
        print(f"Erro ao tentar focar a janela do Excel: {e}")

# Função para colar dados na célula correta da planilha aberta
def colar_dados_na_planilha(linha, data_vencimento, valor_atualizado):
    try:
        focar_janela_excel()  # Assegura que o Excel está em foco

        # Colar data_vencimento na coluna 'R'
        pyautogui.hotkey('ctrl', 'g')  # Abre o "Ir para" no Excel
        time.sleep(1)
        pyautogui.write(f"R{linha + 1}")  # Vai para a célula da coluna 'Data Vencimento Original'
        pyautogui.press('enter')
        time.sleep(1)
        pyperclip.copy(data_vencimento)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)  # Aguarda o tempo necessário para colar
        print(f"Data colada na célula R{linha + 1}.")

        # Colar valor_atualizado na coluna 'Q' usando colagem especial como texto
        pyautogui.hotkey('ctrl', 'g')  # Abre o "Ir para" no Excel
        time.sleep(1)
        pyautogui.write(f"Q{linha + 1}")  # Vai para a célula da coluna 'Valor Atualizado'
        pyautogui.press('enter')
        time.sleep(1)
        pyperclip.copy(valor_atualizado)

        # Menu de Colagem Especial
        pyautogui.hotkey('ctrl', 'alt', 'v')
        time.sleep(1)  # Pequena pausa para garantir que o menu seja aberto
        pyautogui.press('t')  # Seleciona a opção "Texto" no menu de colagem especial
        pyautogui.press('enter')
        time.sleep(1)  # Aguarda o tempo necessário para colar

        print(f"Valor atualizado colado na célula Q{linha + 1}.")
        
    except Exception as e:
        print(f"Erro ao colar dados na planilha: {e}")

# Função principal
def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/Pasta1.xlsx'
    nome_planilha = 'Planilha1'
    tabela_boleto = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, skiprows=3)

    url = 'https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=4&TipoConsultaProcesso=24&ServletRedirect=GuiaFinal_FinalZeroPublica&TituloDaPagina=Guia+Final+e+Final+Zero+%5BAcesso+P%C3%BAblico%5D&hashFluxo=1722628305255'
    
    with webdriver.Chrome(service=Service(ChromeDriverManager().install())) as driver:
        driver.get(url)

        for index, row in tabela_boleto.iterrows():
            copiar_nosso_numero(driver, row['Número Processo'])
            clicar_botao_buscar(driver)
            clicar_emitir_imprimir(driver)
            
            # Extraindo dados da guia usando pyautogui
            data_vencimento, valor_atualizado = extrair_dados_guia()
            if data_vencimento and valor_atualizado:
                colar_dados_na_planilha(index + 4, data_vencimento, valor_atualizado)  # +4 para a linha correta
            
            # Após completar o processamento da linha atual, voltar para a URL original para a próxima iteração
            driver.get(url)

if __name__ == "__main__":
    main()
