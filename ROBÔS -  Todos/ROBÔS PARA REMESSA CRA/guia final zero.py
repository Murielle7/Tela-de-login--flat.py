import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui

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
        botao_buscar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[5]/input[1]'))
        )
        botao_buscar.click()
        print("Botão 'Buscar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'buscar': {e}")

# Função para clicar nos botões "Emitir e Imprimir"
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
        time.sleep(0.5)
    except Exception as e:
        print(f"Erro ao clicar em Emitir e Imprimir: {e}")

# Função para realizar clique em coordenadas específicas
pyautogui.PAUSE = 0.1  # Define um pequeno atraso entre as operações PyAutoGUI
def click(x, y):
    pyautogui.click(x, y)
    time.sleep(0.3)
pyautogui.FAILSAFE = True  # Reativa o fail-safe após a operação

# Função para preencher CPF
def preencher_cpf(driver, cpf, nome):
    campo_cpf = driver.find_element(By.XPATH, '//*[@id="Cpf"]')
    campo_cpf.clear()
    campo_cpf.send_keys(cpf)
    print("Preenchendo CPF...")
    campo_nome_parte = driver.find_element(By.XPATH, '//*[@id="Nome"]')
    campo_nome_parte.clear()
    campo_nome_parte.send_keys(nome)
    print("Nome da parte preenchido.")

# Função para preencher CNPJ
def preencher_cnpj(driver, cnpj, razao_social):
    campo_cnpj = driver.find_element(By.XPATH, '//*[@id="Cnpj"]')
    campo_cnpj.clear()
    campo_cnpj.send_keys(cnpj)
    print("Preenchendo CNPJ...")
    campo_razao_social = driver.find_element(By.XPATH, '//*[@id="RazaoSocial"]')
    campo_razao_social.clear()
    campo_razao_social.send_keys(razao_social)
    print("Razão social preenchida.")

# Função para preencher CPF/CNPJ e outros campos
def preencher_cpf_cnpj(driver, row):
    cpf_cnpj = str(row['CPF/CNPJ']).strip()  # Remove espaços extras
    
    if len(cpf_cnpj) == 11: 
        # Preencher CPF
        checkbox_fisica = driver.find_element(By.XPATH, '//*[@id="tipoPessoaFisica"]')
        if not checkbox_fisica.is_selected():
            checkbox_fisica.click()
        print("Checkbox 'Física' selecionado.")
        preencher_cpf(driver, cpf_cnpj, row['Nome Parte'])
    
    elif len(cpf_cnpj) == 14:  # Garantir que tenha 14 dígitos para CNPJ
        cpf_cnpj = cpf_cnpj.zfill(14)  # Garante que tenha zeros à esquerda para CNPJ
        # Preencher CNPJ
        checkbox_juridica = driver.find_element(By.XPATH, '//*[@id="tipoPessoaJuridica"]')
        if not checkbox_juridica.is_selected():
            checkbox_juridica.click()
        print("Checkbox 'Jurídica' selecionado.")
        preencher_cnpj(driver, cpf_cnpj, row['Nome Parte'])

    # Preencher outros campos
     
    campo_logradouro = driver.find_element(By.XPATH, '//*[@id="Logradouro"]')
    campo_logradouro.clear()
    campo_logradouro.send_keys(row['Logradouro'])
    time.sleep(0.5)
    
    campo_bairro = driver.find_element(By.XPATH, '//*[@id="Bairro"]')
    campo_bairro.clear()
    campo_bairro.send_keys(row['Bairro'])
    time.sleep(0.5)

    campo_cidade = driver.find_element(By.XPATH, '//*[@id="Cidade"]')
    campo_cidade.clear()
    campo_cidade.send_keys(row['Cidade'])
    time.sleep(0.5)

    campo_uf = driver.find_element(By.XPATH, '//*[@id="Uf"]')
    campo_uf.clear()
    campo_uf.send_keys(row['UF'])
    time.sleep(0.5)

    campo_cep = driver.find_element(By.XPATH, '//*[@id="Cep"]')
    campo_cep.clear()
    campo_cep.send_keys(str(row['CEP']))
    time.sleep(0.5)

# Encontrar e clicar no botão "Atualizar"
    botao_atualizar = driver.find_element(By.XPATH, '//*[@id="divBotoesCentralizados"]/button[1]')
    botao_atualizar.click()
    time.sleep(0.5)

# Encontrar e clicar no botão "Emitir e Imprimir"
    botao_emitir_boleto = driver.find_element(By.XPATH, '//*[@id="imgEmitirGuiaImprimirPDF"]')
    botao_emitir_boleto.click()
    time.sleep(0.5)

# Encontrar e clicar no botão "Emitir e Imprimir" final
    botao_emitir_final = driver.find_element(By.XPATH, '/html/body/div[3]/form/div/div[4]/fieldset/div/button')
    botao_emitir_final.click()
    time.sleep(0.5)  # Ajuste o tempo conforme necessário após clicar em "Emitir e Imprimir" final
                
# Função principal
def main():
    caminho_arquivo = r'S:/CAJ/Servidores/Murielle/Dem. PEDRO- Remessa 22 - EMISSÃO DE BOLETO E RENOMEAR ----- P1.xlsx'
    
    tabela_boleto = pd.read_excel(caminho_arquivo, dtype={'CPF/CNPJ': str, 'CEP': str})

    url = 'https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=4&TipoConsultaProcesso=24&ServletRedirect=GuiaFinal_FinalZeroPublica&TituloDaPagina=Guia+Final+e+Final+Zero+%5BAcesso+P%C3%BAblico%5D&hashFluxo'
    driver = webdriver.Chrome()
    driver.get(url)

    for index, row in tabela_boleto.iterrows():
        try:
            # Recarregar a página em vez de abrir uma nova instância do navegador
            driver.get(url)
            copiar_nosso_numero(driver, row['Número Processo'])
            clicar_botao_buscar(driver)
            clicar_emitir_imprimir(driver)
            click(689,570)
            preencher_cpf_cnpj(driver, row)
        except Exception as e:
            print(f"Erro na iteração da linha {index}: {e}")

    driver.quit()

if __name__ == "__main__":
    main()
