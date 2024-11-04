import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def copiar_dados_planilha(origem, destino):
    try:
        df = pd.read_excel(origem, sheet_name='Export', header=0)
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None
    
    colunas_necessarias = ['Número Processo', 'Número Guia', 'Status Débito', 'Nome Parte']  # Adiciona 'Nome Parte'
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print("Erro: uma ou mais colunas não foram encontradas na planilha 'Export'.")
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
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get(url)

        user_field = WebDriverWait(driver, 3).until(
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
        analista_link = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Analista Financeiro')]"))
        )
        analista_link.click()
        print("Clicou em 'Analista Financeiro'.")
    except Exception as e:
        print(f"Erro ao clicar em 'Analista Financeiro': {e}")

def clicar_menu_cadastros(driver):
    try:
        cadastros_menu = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "2"))
        )
        cadastros_menu.click()
        print("Clicou no menu 'Cadastros'.")
        
        financeiros_submenu = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Financeiros')]"))
        )
        financeiros_submenu.click()
        print("Clicou no submenu 'Financeiros'.")

        debitos_submenu = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Débitos - Diretoria Financeira')]"))
        )
        debitos_submenu.click()
        print("Clicou no submenu 'Débitos - Diretoria Financeira'.")
        
    except Exception as e:
        print(f"Erro ao clicar nos menus: {e}")

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
        time.sleep(2)
        processo_input.clear()
        processo_input.send_keys(numero_processo)
        print(f"Colou o número do processo: {numero_processo}")
    except Exception as e:
        print(f"Erro ao colar o número do processo: {e}")

def clicar_botao_consultar(driver):
    try:
        botao_consultar = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/form/div[2]/fieldset/div/input"))
        )
        botao_consultar.click()
        print("Botão 'Consultar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Consultar': {e}")

def localizar_id_e_clicar_editar(driver, nome_parte):
    try:
        # Garantir que cpf_cnpj é uma string
        nome_parte = str(nome_parte)
        
        tabela = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Tabela"]'))
        )
        
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                parte_text = str(coluna.text).strip().replace(".", "").replace("/", "").replace("-", "")
                if parte_text == nome_parte.replace(".", "").replace("/", "").replace("-", ""):
                    print(f"CPF/CNPJ {nome_parte} encontrado na tabela.")
                    botao_editar = linha.find_element(By.XPATH, ".//input[@type='image' and contains(@title, 'Editar Débito')]")
                    botao_editar.click()
                    time.sleep(2)
                    print(f"Clicou em 'Editar Débito' para o CPF/CNPJ: {nome_parte}")
                    return True
        print(f"CPF/CNPJ {nome_parte} não encontrado na tabela.")
        return False
    except Exception as e:
        print(f"Erro ao localizar o CPF/CNPJ e clicar em 'Editar Débito': {e}")
        return False

def selecionar_status(driver, status_processo):
    try:
        select_element = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.ID, "Id_ProcessoDebitoStatus"))
        )

        opcao_em_analise = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/form/div[2]/fieldset[1]/select/option[18]'))
        )
        opcao_em_analise.click()
        print("Selecionou a opção 'PROTESTADO'")
            
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao inserir os dados no campo e realizar as ações: {e}")

def clicar_botao_salvar(driver):
    try:
        botao_salvar = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divPortaBotoes"]/button[1]'))
        )
        botao_salvar.click()
        print("Botão 'Salvar' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Salvar': {e}")

def clicar_botao_confirmar_salvar(driver):
    try:
        botao_confirmar_salvar = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divConfirmarSalvar"]/button'))
        )
        botao_confirmar_salvar.click()
        print("Botão 'Confirmar Salvar' clicado.")
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar no botão 'Confirmar Salvar': {e}")

def clicar_botao_ok(driver):
    try:
        WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[8]/div[3]/div/button'))
        )
        botao_ok = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[3]/div/button'))
        )
        botao_ok.click()
        print("Botão 'Ok' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Ok': {e}")

def clicar_botao_escolher_novo_processo(driver):
    try:
        botao_novo_processo = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/form/div[1]/button[2]'))
        )
        botao_novo_processo.click()
        print("Botão 'Escolher Novo Processo' clicado.")
    except Exception as e:
        print(f"Erro ao clicar no botão 'Escolher Novo Processo': {e}")

def main():
    caminho_arquivo = r'S:/CAJ/Cobrança - Judicial e Administrativa/Custas Judiciais - Finais/Processamento Retorno - CRA/Relatório BI.xlsx'
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
    
    for index, row in df_filtrado.iterrows():
        numero_processo = row['Número Processo']
        status_processo = row['Status Débito']
        nome_parte = row['Nome Parte']  # Recupera o Nome Parte
        
        if pd.isna(nome_parte):  # Verifica se o CPF/CNPJ ou Nome Parte estão vazios
            print("CPF/CNPJ ou Nome Parte vazio encontrado. Finalizando o processamento.")
            break
        
        try:
            copiar_numero_processo(driver, numero_processo)
            clicar_botao_consultar(driver)
            
            if localizar_id_e_clicar_editar(driver, nome_parte):
                selecionar_status(driver, status_processo)
                clicar_botao_salvar(driver)
                clicar_botao_confirmar_salvar(driver)
                clicar_botao_ok(driver)
                df_filtrado.loc[index, 'Alteração de Status'] = "Alterado com sucesso"
                print(f"Processo para {nome_parte} alterado com sucesso.")  # Mensagem informativa
            else:
                df_filtrado.loc[index, 'Alteração de Status'] = "Erro na Alteração"
                print(f"Erro na alteração do processo para {nome_parte}.")  # Mensagem informativa
                clicar_botao_escolher_novo_processo(driver)
                continue
            
            if index != df_filtrado.index[-1]:
                clicar_botao_escolher_novo_processo(driver)
        except Exception as e:
            print(f"Erro ao processar o processo {numero_processo} para {nome_parte}: {e}")
            df_filtrado.loc[index, 'Alteração de Status'] = "Erro na Alteração"

    try:
        df_filtrado.to_excel(caminho_arquivo_destino, index=False)
        print(f"Planilha atualizada com os resultados salvos em: {caminho_arquivo_destino}")
    except Exception as e:
        print(f"Erro ao salvar a planilha atualizada: {e}")
    
    input("Pressione ENTER para fechar o navegador e finalizar o script...")
    driver.quit()

if __name__ == "__main__":
    main()
