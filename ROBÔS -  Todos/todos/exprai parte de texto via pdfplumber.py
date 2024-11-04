from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re


def connect():
    options = Options()
    options.set_preference("dom.webdriver.enabled", False)
    driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()), options=options)
    url = "https://projudi.tjgo.jus.br/GuiaFinal_FinalZeroPublica?PaginaAtual=9"
    driver.get(url)

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
        botao_buscar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div/fieldset/div[5]/input[1]'))
        )
        botao_buscar.click()
        print("Botão 'Buscar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'buscar': {e}")

def busca_guias(fieldsets, processo):

    #print("\nQtd guias: {}".format(len(fieldsets)))
    nrguia = -1
    for fieldset in fieldsets:
        guia =  fieldset.find_element("xpath", "./div[contains(text(), 'Número Guia')]/following-sibling::span").text
        guia = re.sub(r'\D', '', guia)
        #print(f"Guia {guia}")
        nrguia = nrguia + 1
        if re.sub(r'\D', '', guia) == str(int(processo['Número Guia'])):
            data_vencimento_original = fieldset.find_element("xpath", '//div[contains(text(), "Data Vencimento Original")]/following-sibling::span').text
            
            try:
                total_guia_atualizado = fieldset.find_element("xpath", '//td[b[contains(text(), "Total Guia Atualizado")]]/following-sibling::td').text
                clica = fieldset.find_element("xpath", './button[@name="imgEmitirGuiaImprimirPDF"]')
                clica.click()
                WebDriverWait(driver, 5)
                try:
                    clica = fieldset.find_element("xpath", f'./div[@id="divEmitirBoleto{nrguia}"]/fieldset[@id="VisualizaDados"]/div[@id="divBotoesCentralizados{nrguia}"]/button[@name="imgEmitirGuia{nrguia}"]')
                    clica.click()
                    time.sleep(2)
                    
                except:
                    clica = fieldset.find_element("xpath", './div[@id="divEmitirBoleto1"]/fieldset[@id="VisualizaDados"]/div[@id="divBotoesCentralizados1"]/button[@name="imgEmitirGuia1"]')
                    clica.click()
                    time.sleep(2)
                    
                return data_vencimento_original, total_guia_atualizado, True 
                
            except:
                total_guia_atualizado = fieldset.find_element("xpath", '//td[b[contains(text(), "Total da Guia")]]/following-sibling::td').text
                return data_vencimento_original, total_guia_atualizado, False 


def busca_processo(processos, driver):

    df_protestadas = pd.DataFrame()
    df_sem_guias = pd.DataFrame()
    df = pd.DataFrame()
    qtd_processos = processos.shape[0]
    #bar = Bar('Processando', max=qtd_processos)
    for id, processo in processos.iterrows():
        #bar.next()
        ### Apenas processos com status igual a Apto para envio ao protesto
        if processo['Status'] == 'Apto para envio ao protesto':
            #print("Número processo {}".format(processo['Número Processo']))

            clica = driver.find_element("xpath", '//input[@name="imgLimpar"]')
            clica.click()

            WebDriverWait(driver, 5)
            botao_limpa_status = driver.find_element("xpath", '//button[@name="imaLimparProcessoStatus"]')
            # Clicar no botão de login
            botao_limpa_status.click()
            wait = WebDriverWait(driver, 5)
    
            busca = driver.find_element("xpath", '//input[@name="ProcessoNumero"]')
            busca.send_keys(str(processo))
            clica = driver.find_element("xpath", '//input[@name="imgSubmeter"]')
            clica.click()

            try:

                fieldsets = wait.until(EC.visibility_of_all_elements_located(("xpath", '//div[@id="divEditar"]/fieldset[@id="VisualizaDados"]')))
                data_vencimento_original, total_guia_atualizado, status = busca_guias(fieldsets, processo)
                processos.at[id, 'Data Vencimento Original'] = data_vencimento_original
                processos.at[id, 'Valor Atualizado'] = total_guia_atualizado
                
                if status:
                    processos.at[id, 'Observação'] = "OK"                
                else:
                    processos.at[id, 'Observação'] = "Guia sem atualização"

                driver.back()
                
            except:
                try:
                    div_guia = driver.find_element(By.XPATH, '//div[@class="mensagemTexto"]')
                    processos.at[id, 'Observação'] = "Sem guia final"

                except:
                    processos.at[id, 'Observação'] = "Guia Protestada"                

            driver.back()
    
    return processos


#MAIN
driver = connect()

caminho_do_arquivo = r'S:/CAJ/Servidores/Murielle/Teste - dem. Pedro.xlsx'
#caminho_do_arquivo = r'C:\Users\hvnogueira\Documents\relatórios TJGO\python\CAJ\Remessa (1).xlsx'
processos = pd.read_excel(open(caminho_do_arquivo, 'rb'), sheet_name='Planilha1', skiprows=3, header=3)
processos = busca_processo(processos, driver)
processos.to_excel('S:/CAJ/Servidores/Murielle/Teste - dem. Pedro\Planilha1_.xlsx')


print("\n\nO código foi finalizado.\n\n")
input("CLIQUE AQUI e Pressione ENTER para fechar o programa.")
driver.close()
