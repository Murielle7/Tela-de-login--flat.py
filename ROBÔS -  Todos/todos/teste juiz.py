import asyncio
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import pyautogui
import time
import os
import pyperclip

# Função para ler o arquivo Excel e retornar todos os valores da coluna especificada
def get_all_values_from_column(file_path, column_name):
    try:
        df = pd.read_excel(file_path)
        if column_name in df.columns:
            return df[column_name].dropna().tolist()
        else:
            print(f"A coluna '{column_name}' não foi encontrada no arquivo Excel.")
            return []
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return []

# Caminho do arquivo Excel e nome da coluna
file_path = r"\\10.0.20.65\Financeira\CAJ\Automação - Fianca\Restituição - Planilha automação.xlsx"
column_validacao = "Validação Juiz"
column_proad = "Número do PROAD"
column_name_analista = "Nome do Analista"

async def main():
    try:
        # Obtém todos os valores das colunas especificadas do arquivo Excel
        validacoes = get_all_values_from_column(file_path, column_validacao)
        proads = get_all_values_from_column(file_path, column_proad)
        analistas = get_all_values_from_column(file_path, column_name_analista)
    

        # Inicia o Playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            page = await browser.new_page()

            # Itera através de todos os valores para validação
            for value, numero_proad, analista in zip(validacoes, proads, analistas):
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
                        print(f"Criou a pasta '{proad_folder_path}'.")                    # Preenche o campo "Procurar" com o valor da coluna
                    await page.fill('//*[@id="tabela_filter"]/label/input', str(value))
                    print(f"Dados da coluna 'Validação Juiz' inseridos no campo Procurar.")

                    await asyncio.sleep(3)
                    
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
asyncio.run(main())

