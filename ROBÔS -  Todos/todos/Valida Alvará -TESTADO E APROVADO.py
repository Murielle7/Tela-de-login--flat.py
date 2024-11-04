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
file_path = r'\\10.0.20.65\Financeira\CAJ\Automação - Fianca\Restituição - Planilha automação.xlsx'
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

                    # Diretório principal e nova pasta
                    main_dir = r'\\10.0.20.65\Financeira\CAJ\Automação - Fianca'
                    new_folder_path = os.path.join(main_dir, analista)

                    # Cria a nova pasta se não existir
                    if not os.path.exists(new_folder_path):
                        os.makedirs(new_folder_path)
                        print(f"Criou a pasta '{new_folder_path}'.")

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
                    time.sleep(2)  # Tempo para o processo de salvamento
                    print("Clicou no botão 'Enter'.")
                    
                    # Usar pyperclip para copiar o nome do arquivo com o Número do PROAD para a área de transferência
                    nome_arquivo = f"{numero_proad}.pdf"
                    pyperclip.copy(nome_arquivo)
                    pyautogui.hotkey('ctrl', 'v')  # Colar o texto
                    print(f"Nome do arquivo '{nome_arquivo}' definido.")

                    # Navegar para a pasta de destino
                    pyautogui.hotkey('ctrl', 'l')  # Ativar a barra de localização
                    pyautogui.write(new_folder_path)  # Digitar o caminho da nova pasta de destino
                    pyautogui.press('enter')
                    pyautogui.press('enter')
                    print(f"Navegou para a pasta '{new_folder_path}'.")

                    # Pressionar 'Enter' para confirmar o local da pasta
                    pyautogui.press('enter')
                    time.sleep(2)

                    # Pressionar 'Enter' novamente para confirmar o salvamento
                    pyautogui.press('enter')
                    time.sleep(2)  # Tempo para o processo de salvamento
                    print("Clicou duas vezes no botão 'Enter'.")
                    
                except Exception as e:
                    print(f"Erro durante a automação: {e}")
                finally:
                    await browser.close()
                    time.sleep(2)  # Aguarda um pouco antes de começar a próxima iteração

    except Exception as e:
        print(f"Erro ao iniciar a automação: {e}")

# Executa a função principal
asyncio.run(main())
