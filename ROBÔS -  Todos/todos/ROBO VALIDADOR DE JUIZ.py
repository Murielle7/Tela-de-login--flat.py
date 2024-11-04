import asyncio
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import time
import os
import openpyxl  # Biblioteca para manipulação do Excel
import pyautogui
import pyperclip

# Função para ler o arquivo Excel e retornar o DataFrame
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return pd.DataFrame()

# Função para escrever o DataFrame de volta no arquivo Excel
def write_excel(df, file_path):
    try:
        df.to_excel(file_path, index=False)
        print(f"Arquivo Excel atualizado: {file_path}")
    except Exception as e:
        print(f"Erro ao escrever no arquivo Excel: {e}")

# Caminho do arquivo Excel e nome das colunas
file_path = r"\\10.0.20.65\Financeira\CAJ\Automação - Fianca\Restituição - Planilha automação.xlsx"
column_validacao = "Validação Juiz"
column_proad = "Número do PROAD"
column_name_analista = "Nome do Analista"
column_status = "Status Juíz"

async def main(df):
    try:
        # Inicia o Playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            page = await browser.new_page()

            # Itera através dos valores das colunas especificadas
            for index, row in df.iterrows():
                value = row[column_validacao]
                numero_proad = row[column_proad]
                analista = row[column_name_analista]
                
                try:
                    await page.goto('https://docs.tjgo.jus.br/comarcas/foruns/listaJuizes.html')
                    time.sleep(2)

                    # Diretório principal e nova pasta por analista
                    main_dir = r"\\10.0.20.65\Financeira\CAJ\Automação - Fianca"
                    analista_folder_path = os.path.join(main_dir, analista)
                    
                    # Cria a pasta do analista, se não existir
                    if not os.path.exists(analista_folder_path):
                        os.makedirs(analista_folder_path)
                        print(f"Criou a pasta '{analista_folder_path}'.")

                    # Cria subpasta para o Número do PROAD do analista 
                    proad_folder_path = os.path.join(analista_folder_path, str(numero_proad))
                    
                    if not os.path.exists(proad_folder_path):
                        os.makedirs(proad_folder_path)
                        print(f"Criou a pasta '{proad_folder_path}'.")

                    # Preenche o campo "Procurar" com o valor da coluna
                    await page.fill('//*[@id="tabela_filter"]/label/input', str(value))
                    print(f"Dados da coluna 'Validação Juiz' inseridos no campo Procurar.")

                    await asyncio.sleep(3)
                    
                    # Captura o conteúdo HTML da página
                    page_content = await page.content()
                    
                    # Verifica se a mensagem "Nenhum registro localizado" está presente na página
                    if "Nenhum registro localizado" in page_content:
                        status = "ERRO"
                    else:
                        status = "FEITO"
                    
                    # Atualizando o DataFrame com o status
                    df.at[index, column_status] = status
                    
                    # Salve o status atualizado de volta no arquivo Excel
                    write_excel(df, file_path)

                    # Simular "Ctrl+P" para abrir a janela de impressão
                    pyautogui.hotkey('ctrl', 'p')
                    time.sleep(2)  # Espera a janela de impressão abrir
                    print("'Deu Ctrl + P'.")

                    # Pressionar 'Enter' para confirmar o salvamento
                    pyautogui.press('enter')
                    time.sleep(2)

                    # Copiar o caminho completo do arquivo para a área de transferência
                    nome_arquivo = f"Validação de Juiz"
                    caminho_arquivo = os.path.join(proad_folder_path, nome_arquivo)
                    pyperclip.copy(caminho_arquivo)
                    pyautogui.hotkey('ctrl', 'a')  # Selecionar a barra de localização do nome do arquivo
                    pyautogui.hotkey('ctrl', 'v') 
                    print(f"Caminho do arquivo '{caminho_arquivo}' definido.")

                    # Pressionar 'Enter' para confirmar o local do arquivo
                    pyautogui.press('enter')
                    time.sleep(2)  # Tempo para o processo de salvamento
                    print("Salvou o arquivo PDF.")

                    # Navegar de volta para a página principal
                    await page.goto('https://docs.tjgo.jus.br/comarcas/foruns/listaJuizes.html')
                    print("Navegou de volta à página principal.")
                except Exception as e:
                    print(f"Ocorreu um erro ao processar o número PROAD {numero_proad} com a validação {value}: {e}")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Execução do script principal
df = read_excel(file_path)
asyncio.run(main(df))
