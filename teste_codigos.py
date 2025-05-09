import os
import pandas as pd
import pyautogui
import time
import subprocess
from pywinauto import Application

# Caminho para a pasta de downloads
download_folder = os.path.expanduser('G:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\teste_auto_455')

# Função para encontrar o último arquivo .sswweb
def encontrar_ultimo_arquivo_swwweb(pasta):
    arquivos = [os.path.join(pasta, f) for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
    arquivos = [f for f in arquivos if f.endswith('.sswweb')]
    return max(arquivos, key=os.path.getctime) if arquivos else None

# Função para processar o arquivo .sswweb
def processar_arquivo_swwweb(download_folder, planilha_destino):
    ultimo_arquivo = encontrar_ultimo_arquivo_swwweb(download_folder)

    if ultimo_arquivo:
        print(f"Último arquivo encontrado: {ultimo_arquivo}")

        # Lê o conteúdo do arquivo e remove as duas primeiras linhas
        with open(ultimo_arquivo, 'r') as file:
            linhas = file.readlines()

        # Remove as duas primeiras linhas
        linhas = linhas[2:]

        # Grava o conteúdo de volta no arquivo
        with open(ultimo_arquivo, 'w') as file:
            file.writelines(linhas)

        # Abre o arquivo no Bloco de Notas
        subprocess.Popen(['notepad.exe', ultimo_arquivo])  # Força a abertura no Bloco de Notas

        # Aguarda um tempo para garantir que o arquivo seja aberto
        time.sleep(5)  # Ajuste o tempo conforme necessário

        # Copia todos os dados restantes do Bloco de Notas
        pyautogui.hotkey('ctrl', 'a')  # Seleciona tudo
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')  # Copia os dados

        # Aguarda um tempo para garantir que os dados sejam copiados
        time.sleep(1)

        # Abre a planilha de destino no Excel
        if not os.path.exists(planilha_destino):
            # Cria um DataFrame vazio e salva como um novo arquivo Excel
            pd.DataFrame().to_excel(planilha_destino, index=False)
            print(f"Planilha criada: {planilha_destino}")

        os.startfile(planilha_destino)
        time.sleep(10)  # Aumenta o tempo de espera para garantir que o Excel esteja aberto

        # Foca na janela do Excel usando pywinauto
        app = Application().connect(title_re=".*Excel.*")  # Conecta à janela do Excel
        excel_window = app.top_window()  # Obtém a janela principal do Excel
        excel_window.set_focus()  # Define o foco na janela do Excel

        # Aguarda um tempo para garantir que o Excel esteja em foco
        time.sleep(1)

        # Lê a planilha de destino
        planilha_destino_dados = pd.read_excel(planilha_destino)

        # Encontra a última linha preenchida na planilha de destino
        ultima_linha = len(planilha_destino_dados)

        # Clica na primeira coluna da próxima linha
        excel_window.type_keys(f"A{ultima_linha + 1}", with_spaces=True)  # Move para a célula da primeira coluna da próxima linha
        time.sleep(1)  # Aguarda um momento para garantir que a célula esteja focada

        # Cola os dados copiados na célula selecionada
        pyautogui.hotkey('ctrl', 'v')  # Cola os dados

        # Aguarda um tempo para garantir que os dados sejam colados
        time.sleep(1)

        # Salva os dados na planilha de destino
        planilha_destino_dados.to_excel(planilha_destino, index=False)
        print(f"Dados copiados para a planilha: {planilha_destino}")
    else:
        print("Nenhum arquivo .sswweb encontrado na pasta de downloads.")

# Caminho para a planilha de destino
planilha_destino = os.path.join(download_folder, "dados_copiados.xlsx")

# Chama a função para processar o arquivo
processar_arquivo_swwweb(download_folder, planilha_destino)