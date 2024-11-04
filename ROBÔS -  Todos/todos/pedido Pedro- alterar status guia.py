import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Função para conectar ao WebDriver
def connect():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Executa em modo headless (sem interface gráfica), remova se quiser ver a UI
    driver = webdriver.Chrome(options=options)
    return driver

# Função para copiar o número do processo em um campo
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
        processo_input.clear()
        processo_input.send_keys(numero_processo)
        print(f"Colou o número do processo: {numero_processo}")
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
        botao_emitir_boleto = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="imgEmitirGuiaImprimirPDF"]'))
        )
        botao_emitir_boleto.click()

        botao_emitir_final = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[9]/fieldset/div/button'))
        )
        botao_emitir_final.click()
    except Exception as e:
        print(f"Erro ao clicar em Emitir e Imprimir: {e}")

# Função auxiliar para depuração de DataFrame
def debug_dataframe(df):
    print("Cabeçalho das colunas:", df.columns)
    print("Algumas linhas do DataFrame:")
    print(df.head())

# Função para processar os números de processo
def numero_processo(df, driver):
    for index, row in df.iterrows():
        process_number = row['Processo']  # Atualize com o nome correto da coluna de processo
        copiar_numero_processo(driver, process_number)
        clicar_botao_buscar(driver)
        clicar_emitir_imprimir(driver)
    return df

# Função principal
def main():
    driver = connect()
    try:
        caminho_do_arquivo = r'S:\CAJ\Cobrança - Judicial e Administrativa\Custas Judiciais - Finais\Remessas - CRA\Remessa.xlsx'
        processos = pd.read_excel(open(caminho_do_arquivo, 'rb'), sheet_name='Conferência - GO', skiprows=3, header=0)
        
        # Processando os números dos processos
        processos = numero_processo(processos, driver)
        
        # Exportando o DataFrame modificado para um novo arquivo Excel
        processos.to_excel(r'S:\CAJ\Cobrança - Judicial e Administrativa\Custas Judiciais - Finais\Remessas - CRA\Conferencia_.xlsx')

        print("\n\nO código foi finalizado.\n\n")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        input("CLIQUE AQUI e Pressione ENTER para fechar o programa.")
        driver.quit()  # Use quit() para fechar o WebDriver corretamente

if __name__ == "__main__":
    main()
