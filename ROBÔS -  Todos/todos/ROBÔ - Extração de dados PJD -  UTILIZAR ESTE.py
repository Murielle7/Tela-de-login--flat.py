import pyautogui
import pyperclip
import pygetwindow as gw
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

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
        # Localiza a janela do navegador por parte do título
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
        # Simula Ctrl+F no navegador para abrir a caixa de pesquisa
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        
        # Selecionar todo o texto atual
        pyautogui.hotkey('ctrl', 'a')   
        time.sleep(0.5)  # Pequena pausa para garantir que a seleção foi feita
        
        # Apagar o texto selecionado
        pyautogui.press('backspace')   
        time.sleep(0.5)  # Dê tempo para que o campo de texto seja atualizado
        
        # Insere o novo número da guia como string
        pyautogui.write(str(numero_guia))
        pyautogui.press('enter')  # Confirmação da busca com Enter
        print(f"Número da guia '{numero_guia}' digitado no navegador.")
    except Exception as e:
        print(f"Erro ao digitar número guia: {e}")

def localizar_e_copiar_data(driver):
    try:
        # Lista de seletores CSS a tentar
        seletores_css = [
            "fieldset.VisualizaDados:nth-child(4) > span:nth-child(18)",
            "div.divEditar:nth-child(16) > fieldset:nth-child(1) > span:nth-child(18)"
        ]

        data_element = None
        index = 0

        # Tenta encontrar o elemento com diferentes seletores CSS
        while data_element is None and index < len(seletores_css):
            try:
                data_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, seletores_css[index]))
                )
            except Exception:
                data_element = None
                index += 1

        if not data_element:
            raise Exception("Elemento de data não encontrado com os seletores fornecidos.")

        # Calcula as coordenadas do elemento
        location = data_element.location
        size = data_element.size
        x_mid = location['x'] + size['width'] // 2
        y_mid = location['y'] + size['height'] // 2
        
        # certificado de que a janela do navegador está em foco
        driver.switch_to.window(driver.current_window_handle)
        
        pyautogui.moveTo(x_mid, y_mid)
        pyautogui.click(clicks=3)  # Triplo clique para selecionar a data
        pyautogui.hotkey('ctrl', 'c')  # Copiar a data
        data_vencimento_original = data_element.text.strip()

        print(f"Data de vencimento copiada: {data_vencimento_original}")
        return data_vencimento_original
    except Exception as e:
        print(f"Erro ao copiar a data de vencimento: {e}")
        return None
    
def clicar_emitir_imprimir(driver):
    try:
        botao_emitir_boleto = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="imgEmitirGuiaImprimirPDF"]'))
        )
        botao_emitir_boleto.click()

        botao_emitir_final = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[9]/fieldset/div/button'))
        )
        botao_emitir_final.click()
        time.sleep(1)  # Espera para garantir que a página seja carregada
        print("Botões 'Emitir' e 'Imprimir' clicados.")
    except Exception as e:
        print(f"Erro ao clicar nos botões Emitir e Imprimir: {e}")

def extrair_valor_atualizado():
    try:
        x_start, y_start = 982, 651  # Ponto inicial da seleção
        x_end, y_end = 1049, 651     # Ponto final da seleção
        
        # Move o cursor para a posição inicial e faz a seleção
        pyautogui.moveTo(x_start, y_start, duration=0.5)
        pyautogui.mouseDown()  # Pressiona o botão do mouse 
        time.sleep(0.5)  # Tempo para o pressionar ser registrado
        pyautogui.moveTo(x_end, y_end, duration=0.5)  # Arrasta até o ponto final
        pyautogui.mouseUp()  # Solta o botão do mouse

        time.sleep(0.5)  # Espera antes de copiar
        pyautogui.hotkey('ctrl', 'c')  # Copiar o texto selecionado
        time.sleep(0.5)  # Tempo para garantir a cópia

        valor_atualizado = pyperclip.paste().strip()  # Pega o texto da área de transferência
        print(f"Valor atualizado extraído: {valor_atualizado}")

        return valor_atualizado
    except Exception as e:
        print(f"Erro ao extrair valor atualizado: {e}")
        return None

def focar_janela_excel():
    try:
        # Localiza a janela do Excel por título
        excel_windows = [win for win in gw.getAllTitles() if 'Excel' in win]
        
        if excel_windows:
            # Seleciona a primeira janela encontrada
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
        focar_janela_excel()

        # Ir para a célula especificada e colar o valor
        pyautogui.hotkey('ctrl', 'g')  # Abre o "Ir para" no Excel
        time.sleep(1)
        pyautogui.write(f"{coluna}{linha + 1}")
        pyautogui.press('enter')
        time.sleep(1)
        
        pyperclip.copy(valor)  # Copiar valor para a área de transferência
        pyautogui.hotkey('ctrl', 'v')  # Colar valor
        print(f"Valor '{valor}' colado na célula {coluna}{linha + 1}.")
        
    except Exception as e:
        print(f"Erro ao colar dados na planilha: {e}")

def salvar_planilha():
    try:
        focar_janela_excel()
        pyautogui.hotkey('ctrl', 's')  # Salvar a planilha
        time.sleep(1)
       
        print("Planilha salva.")
    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")

def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/Pasta1.xlsx'
    nome_planilha = 'Planilha2'
    tabela_boleto = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, skiprows=3)

    url = 'https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=4&TipoConsultaProcesso=24&ServletRedirect=GuiaFinal_FinalZeroPublica&TituloDaPagina=Guia+Final+e+Final+Zero+%5BAcesso+P%C3%BAblico%5D&hashFluxo=1722628305255'
    
    options = webdriver.FirefoxOptions()
    with webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options) as driver:
        acessar_url(driver, url)

        for index, row in tabela_boleto.iterrows():
            nosso_numero = row['Número Processo']
            numero_guia = row['Número Guia']  # Supondo que esse é o nome da coluna com o número da guia
            
            copiar_nosso_numero(driver, nosso_numero)
            clicar_botao_buscar(driver)

            # Foque no navegador antes de simular Ctrl+F
            focar_janela_navegador("Projudi")  # Passando parte do título esperado do navegador
            buscar_numero_guia(driver, numero_guia)

            # Extração e colagem da data de vencimento
            data_vencimento_original = localizar_e_copiar_data(driver)
            if data_vencimento_original:
                colar_dados_na_planilha(4 + index, 'R', data_vencimento_original)

            # Clicar nos botões Emitir e Imprimir
            clicar_emitir_imprimir(driver)

            # Extração e colagem do valor atualizado
            valor_atualizado = extrair_valor_atualizado()
            if valor_atualizado:
                colar_dados_na_planilha(4 + index, 'Q', valor_atualizado)

            # Salvar a planilha
            salvar_planilha()

            print("Reiniciando para o próximo número do processo...\n")
            time.sleep(1)

            # Voltar para a URL original para a próxima iteração
            acessar_url(driver, url)

if __name__ == "__main__":
    main()
