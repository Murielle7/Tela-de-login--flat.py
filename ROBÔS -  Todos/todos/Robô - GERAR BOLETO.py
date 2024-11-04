import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

# Defina o caminho do arquivo
caminho_arquivo = r'S:/CAJ/Servidores/Murielle/demanda giseleplanilha 3.xlsx'

# Função para formatar CPF/CNPJ mantendo os zeros à esquerda
def formatar_cpf_cnpj(numero):
    if len(numero) <= 10:
        return numero.zfill(11)  # Preencher com zeros à esquerda até completar 11 dígitos
    elif len(numero) >= 12:
        return numero.zfill(14)  # Preencher com zeros à esquerda até completar 14 dígitos para CNPJ
    return numero
def adicionar_zero_cnpj(cnpj):
    if len(cnpj) == 10:  # Se o CNPJ possui 13 dígitos
        cnpj = '0' + cpf  # Adiciona zero à esquerda
    return cpf
def adicionar_zero_cnpj(cnpj):
    if len(cnpj) == 12:  # Se o CNPJ possui 13 dígitos
        cnpj = '0' + cnpj  # Adiciona zero à esquerda
    return cnpj
def adicionar_zero_cnpj(cnpj):
    if len(cnpj) == 13:  # Se o CNPJ possui 13 dígitos
        cnpj = '0' + cnpj  # Adiciona zero à esquerda
    return cnpj

# Função auxiliar para depuração de DataFrame
def debug_dataframe(df):
    print("Cabeçalho das colunas:", df.columns)
    print("Algumas linhas do DataFrame:")
    print(df.head())

# Leia a planilha 'Planilha1' do arquivo Excel a partir da linha 4
try:
    df = pd.read_excel(caminho_arquivo, sheet_name='Planilha1', engine='openpyxl', header=0)  # Inicia na linha 1 (índice 0)
    print("Planilha 'Planilha1' lida com sucesso a partir da linha 1!")
    debug_dataframe(df)  # Debug: Mostrar informações sobre o DataFrame lido
    
    # Automatizar interação com o navegador usando Selenium
    driver = webdriver.Chrome()  # Certifique-se de ter o ChromeDriver configurado e acessível
    
    for index, row in df.iterrows():
        if row['Status'] == 'Apto para envio ao protesto':
            numero_guia = str(int(row['NUMEROTITULO']))
            cpf_cnpj = str(int(row['CPFCNPJ']))  # Convertendo CPF/CNPJ para string, mantendo a quantidade de dígitos
            nome_parte = str(row.get('DEVEDOR', ''))
            logradouro = str(row.get('Endereço', ''))
            bairro = str(row.get('Bairro', ''))
            cidade = str(row.get('Cidade', ''))
            uf = str(row.get('UF', ''))
            cep = str(row.get('CEP', ''))
            print(f"Status: Apto para envio ao protesto - Número Guia: {numero_guia}")

            if cpf_cnpj.lower() == 'generico':  # Condição para CPF genérico (altere 'generico' conforme sua necessidade)
                print("CPF genérico encontrado, pulando para a próxima linha.")
                continue  # Pular para a próxima iteração caso seja um CPF genérico
            
            cpf_cnpj_formatado = formatar_cpf_cnpj(cpf_cnpj)
            cpf_cnpj_formatado = adicionar_zero_cnpj(cpf_cnpj_formatado)
            
            # Navegar até a página de emissão de boleto
            driver.get('https://projudi.tjgo.jus.br/GerarBoleto?PaginaAtual=4')

            # Preencher o campo Número Guia
            campo_numero_guia = driver.find_element(By.XPATH, '//*[@id="numeroGuiaConsulta"]')  # Localizar o elemento pelo xpath
            campo_numero_guia.clear()  # Limpar o campo antes de preencher com um novo número de guia
            campo_numero_guia.send_keys(numero_guia)
            time.sleep(1)
            
            # Encontrar e clicar no botão "Gerar Boleto"
            botao_gerar_boleto = driver.find_element(By.XPATH, '//*[@id="divBotoesCentralizados"]/button[1]')
            botao_gerar_boleto.click()
            time.sleep(1)
            
            cpf_cnpj_formatado = formatar_cpf_cnpj(cpf_cnpj)

            # Verificar se é CPF ou CNPJ e marcar o checkbox correspondente

            # Pessoa FÍSICA
            if len(cpf_cnpj_formatado) == 11:  # Se tiver 11 dígitos, considera-se CPF
                checkbox_fisica = driver.find_element(By.XPATH, '//*[@id="tipoPessoaFisica"]')
                if not checkbox_fisica.is_selected():
                    checkbox_fisica.click()
                print("Checkbox 'Física' selecionado.")
                if len(cpf_cnpj_formatado) == 11:  # Se for CPF
                    campo_cpf = driver.find_element(By.XPATH, '//*[@id="Cpf"]')
                    campo_cpf.clear()
                    campo_cpf.send_keys(cpf_cnpj_formatado)
                    time.sleep(1)
                    campo_nome_parte = driver.find_element(By.XPATH, '//*[@id="Nome"]') 
                    campo_nome_parte.clear()
                    campo_nome_parte.send_keys(nome_parte)
            print(f"Nome '{nome_parte}' inserido com sucesso no campo 'Nome' do site.")

            # Pessoa JURÍDICA
            if len(cpf_cnpj_formatado) >11:  # Se for CNPJ
                checkbox_juridica = driver.find_element(By.XPATH, '//*[@id="tipoPessoaJuridica"]')
                if not checkbox_juridica.is_selected():
                    checkbox_juridica.click()
                print("Checkbox 'Jurídica' selecionado.")
            if len(cpf_cnpj_formatado) > 11:  # Se for CNPJ
                campo_cnpj = driver.find_element(By.XPATH, '//*[@id="Cnpj"]')
                campo_cnpj.clear()
                campo_cnpj.send_keys(cpf_cnpj_formatado)
                time.sleep(1)
                campo_razao_social = driver.find_element(By.XPATH, '//*[@id="RazaoSocial"]')
                campo_razao_social.clear()
                campo_razao_social.send_keys(nome_parte)
            print(f"Razão Social '{nome_parte}' inserido com sucesso no campo 'Razão Social' do site.")

            # Campo Endereço
            campo_logradouro = driver.find_element(By.XPATH, '//*[@id="Logradouro"]')
            logradouro = row['Endereço'] # Obtém o valor da coluna 'Endereço Formatado' da linha atual (row)
            campo_logradouro.clear()  # Limpar o campo antes de preencher com novos dados
            campo_logradouro.send_keys(logradouro)
            print(f"Endereço '{logradouro}' inserido com sucesso no campo 'Logradouro' do site.")

            # Preencher Bairro
            campo_bairro = driver.find_element(By.XPATH, '//*[@id="Bairro"]')
            bairro = row['Bairro'] # Obtém o valor da coluna 'Bairro' da linha atual (row)
            campo_bairro.clear()  # Limpar o campo antes de preencher com novos dados
            campo_bairro.send_keys(bairro)
            print(f"Bairro '{bairro}' inserido com sucesso no campo 'Bairro' do site.")

            # Preencher Cidade
            campo_cidade = driver.find_element(By.XPATH, '//*[@id="Cidade"]')
            cidade = row['Cidade'] # Obtém o valor da coluna 'Cidade' da linha atual (row)
            campo_cidade.clear()  # Limpar o campo antes de preencher com novos dados           
            campo_cidade.send_keys(cidade)
            print(f"Cidade '{cidade}' inserida com sucesso no campo 'Cidade' do site.")

            # Preencher UF
            campo_uf = driver.find_element(By.XPATH, '//*[@id="Uf"]')
            uf = row['UF']  # Obtém o valor da coluna 'UF' da linha atual (row)
            campo_uf.clear()  # Limpar o campo antes de preencher com novos dados
            campo_uf.send_keys(uf)
            print(f"UF '{uf}' inserida com sucesso no campo 'UF' do site.")

            # Preencher CEP
            campo_cep = driver.find_element(By.XPATH, '//*[@id="Cep"]')
            cep = str(row['CEP'])  # Obtém o valor da coluna 'CEP' da linha atual (row)
            campo_cep.clear()  # Limpar o campo antes de preencher com novos dados
            campo_cep.send_keys(cep)
            print(f"CEP '{cep}' inserido com sucesso no campo 'CEP' do site.")

            # Encontrar e clicar no botão "Atualizar"
            botao_atualizar = driver.find_element(By.XPATH, '//*[@id="divBotoesCentralizados"]/button[1]')
            botao_atualizar.click()

            # Encontrar e clicar no botão "Emitir e Imprimir"
            botao_emitir_boleto = driver.find_element(By.XPATH, '//*[@id="imgEmitirGuiaImprimirPDF"]')
            botao_emitir_boleto.click()
            
            # Encontrar e clicar no botão "Emitir e Imprimir" final botao_emitir_final=driver.find_element(By.XPATH,'/html/body/div[3]/form/div/div[4]/fieldset/div/button')
            botao_emitir_boleto_final = driver.find_element(By.XPATH, '/html/body/div[3]/form/div/div[4]/fieldset/div/button')
            botao_emitir_boleto_final.click()

    # Fechar o navegador após o término do processamento
    driver.quit()

except Exception as e:
    print (f"Ocorreu um erro ao ler a planilha 'Planilha1': {e}")
