import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import pyautogui

# Função para formatar CPF/CNPJ com zeros à esquerda
def formatar_cpf_cnpj(valor):
    num_digits = len(valor)
    if num_digits <= 11:
        return valor.zfill(11)  # CPF com 11 dígitos
    elif num_digits <= 14:
        return valor.zfill(14)  # CNPJ com 14 dígitos
    else:
        return valor

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

        processo_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ProcessoNumero"))
        )
        time.sleep(5)
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
        botao_buscar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"][value="Buscar"]'))
        )
        botao_buscar.click()
        print("Botão 'Buscar' clicado.")
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao clicar no botão 'Buscar': {e}")

# Função para clicar nos botões "Emitir e Imprimir"
def clicar_emitir_imprimir(driver):
    try:
        botao_emitir_boleto = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/button'))
        )
        botao_emitir_boleto.click()
        time.sleep(5)

        botao_emitir_final = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[9]/fieldset/div/button')) 
        )
        botao_emitir_final.click()
        time.sleep(2)
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao clicar em Emitir e Imprimir: {e}")

# Função para realizar clique em coordenadas específicas
def click(x, y):
    pyautogui.click(x, y)
    time.sleep(0.1)

# Função para preencher CPF/CNPJ e outros campos
def preencher_cpf_cnpj(driver, row):
    try:
        cpf_cnpj = str(row['CPF/CNPJ'])
        cpf_cnpj = formatar_cpf_cnpj(cpf_cnpj)  # Formatar CPF/CNPJ com zeros à esquerda

        if len(cpf_cnpj) == 11:
            checkbox_fisica = driver.find_element(By.ID, 'tipoPessoaFisica')
            if not checkbox_fisica.is_selected():
                checkbox_fisica.click()
            print("Checkbox 'Física' selecionado.")
            campo_cpf = driver.find_element(By.ID, 'Cpf')
            campo_cpf.clear()
            campo_cpf.send_keys(cpf_cnpj)
            time.sleep(2)
            print("Preenchendo CPF...")
            campo_nome_parte = driver.find_element(By.ID, 'Nome')
            campo_nome_parte.clear()
            campo_nome_parte.send_keys(row['Nome Parte'])
            time.sleep(1)

        elif len(cpf_cnpj) == 14:
            checkbox_juridica = driver.find_element(By.ID, 'tipoPessoaJuridica')
            if not checkbox_juridica.is_selected():
                checkbox_juridica.click()
            print("Checkbox 'Jurídica' selecionado.")
            campo_cnpj = driver.find_element(By.ID, 'Cnpj')
            campo_cnpj.clear()
            campo_cnpj.send_keys(cpf_cnpj)
            time.sleep(2)
            print("Preenchendo CNPJ...")
            campo_razao_social = driver.find_element(By.ID, 'RazaoSocial')
            campo_razao_social.clear()
            campo_razao_social.send_keys(row['Nome Parte'])
            time.sleep(1)

        campo_logradouro = driver.find_element(By.ID, 'Logradouro')
        campo_logradouro.clear()
        campo_logradouro.send_keys(row['Logradouro'])
        time.sleep(1)

        campo_bairro = driver.find_element(By.ID, 'Bairro')
        campo_bairro.clear()
        campo_bairro.send_keys(row['Bairro'])
        time.sleep(1)

        campo_cidade = driver.find_element(By.ID, 'Cidade')
        campo_cidade.clear()
        campo_cidade.send_keys(row['Cidade'])
        time.sleep(1)

        campo_uf = driver.find_element(By.ID, 'Uf')
        campo_uf.clear()
        campo_uf.send_keys(row['UF'])
        time.sleep(1)

        campo_cep = driver.find_element(By.ID, 'Cep')
        campo_cep.clear()
        campo_cep.send_keys(str(row['CEP']))
        time.sleep(1)

        # Clicar em "Atualizar"
        botao_atualizar = driver.find_element(By.XPATH, '//*[@id="divBotoesCentralizados"]/button[1]')
        botao_atualizar.click()
        time.sleep(1)

        # Clicar em "Emitir e Imprimir"
        botao_emitir_boleto = driver.find_element(By.ID, 'imgEmitirGuiaImprimirPDF')
        botao_emitir_boleto.click()
        time.sleep(1)

        # Clicar em "Emitir e Imprimir" final
        botao_emitir_final = driver.find_element(By.XPATH, '/html/body/div[3]/form/div/div[4]/fieldset/div/button')
        botao_emitir_final.click()
        time.sleep(2)

    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao preencher campos: {e}")

# Função principal
def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/1.xlsx'
    
    tabela_boleto = pd.read_excel(caminho_arquivo)

    url = 'https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=4&TipoConsultaProcesso=24&ServletRedirect=GuiaFinal_FinalZeroPublica&TituloDaPagina=Guia+Final+e+Final+Zero+%5BAcesso+P%C3%BAblico%5D&hashFluxo'
    driver = webdriver.Chrome()
    driver.get(url)

    for index, row in tabela_boleto.iterrows():
        try:
            copiar_numero_processo(driver, row['Número Processo'])
            clicar_botao_buscar(driver)
            clicar_emitir_imprimir(driver)
            click(700, 568)
            preencher_cpf_cnpj(driver, row)
            driver.get(url)  # Voltar para a URL original para a próxima iteração
        except Exception as e:
            print(f"Erro durante o processamento da linha {index}: {e}")

    driver.quit()

if __name__ == "__main__":
    main()
