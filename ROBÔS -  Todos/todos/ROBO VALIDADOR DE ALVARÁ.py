import asyncio
from playwright.async_api import async_playwright
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

# Caminho do arquivo Excel e nome das colunas
file_path = r"\\10.0.20.65\Financeira\CAJ\Automação - Fianca\Restituição - Planilha automação.xlsx"
column_name_validacao = "Validação Alvará"
column_name_proad = "Número do PROAD"
column_name_analista = "Nome do Analista"

async def main():
    try:
        # Obtém todos os valores das colunas especificadas do arquivo Excel
        validacoes = get_all_values_from_column(file_path, column_name_validacao)
        proads = get_all_values_from_column(file_path, column_name_proad)
        analistas = get_all_values_from_column(file_path, column_name_analista)

        # Itera através de todos os valores para validação
        for value, numero_proad, analista in zip(validacoes, proads, analistas):
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False)
                page = await browser.new_page()
                try:
                    await page.goto('https://projudi.tjgo.jus.br/PendenciaPublica')

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

                    # Preenche o campo "Código Localizador" com o valor da coluna
                    await page.fill('//*[@id="codPublicacao"]', str(value))
                    print(f"Dados da coluna 'Validação Alvará' inseridos no campo codPublicacao.")

                    # Aguardando um pouco para visualização
                    await asyncio.sleep(3)

                    # Clicar no botão "Localizar"
                    await page.click('//*[@id="formLocalizar"]/button')
                    print("Clicou no botão 'Localizar'.")

                    await asyncio.sleep(3)
                    
                    # Simular "Ctrl+P" para abrir a janela de impressão
                    pyautogui.hotkey('ctrl', 'p')
                    time.sleep(2)  # Espera a janela de impressão abrir
                    print("'Deu Ctrl + P'.")

                    # Pressionar 'Enter' para confirmar o salvamento
                    pyautogui.press('enter')
                    time.sleep(2)

                    # Copiar o caminho completo do arquivo para a área de transferência
                    nome_arquivo = f"Validação de Alvará.pdf"
                    caminho_arquivo = os.path.join(proad_folder_path, nome_arquivo)
                    pyperclip.copy(caminho_arquivo)
                    pyautogui.hotkey('ctrl', 'a')  # Selecionar a barra de localização do nome do arquivo
                    pyautogui.hotkey('ctrl', 'v') 
                    print(f"Caminho do arquivo '{caminho_arquivo}' definido.")

                    # Pressionar 'Enter' para confirmar o local do arquivo
                    pyautogui.press('enter')
                    time.sleep(2)  # Tempo para o processo de salvamento
                    print("Salvou o arquivo PDF.")
                    
                except Exception as e:
                    print(f"Erro durante a automação: {e}")
                finally:
                    await browser.close()
                    time.sleep(2)  # Aguarda um pouco antes de começar a próxima iteração

    except Exception as e:
        print(f"Erro ao iniciar a automação: {e}")

# Executa a função principal
asyncio.run(main())
