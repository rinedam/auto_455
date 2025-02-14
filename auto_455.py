from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
import os
import time

# Caminho para a pasta de downloads desejada
download_folder = os.path.expanduser('G:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\teste_auto_455')

# Configurações do Edge
edge_options = Options()
edge_options.add_experimental_option('prefs', {
    "download.default_directory": download_folder,  # Define o diretório de download
    "download.prompt_for_download": False,           # Não solicitar confirmação para download
    "download.directory_upgrade": True,               # Permitir a atualização do diretório
    "safebrowsing.enabled": True                       # Habilitar navegação segura
})

# Inicializa o WebDriver
driver = webdriver.Edge(options=edge_options)

try:
    # Navega até a página do formulário
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")  # Substitua pela URL do seu formulário

    # Atraso para garantir que a página carregue completamente
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))

    # Preenche os campos de login
    driver.find_element(By.NAME, "f1").send_keys("LDI")
    driver.find_element(By.NAME, "f2").send_keys("12373493977")  # Exemplo de CPF
    driver.find_element(By.NAME, "f3").send_keys("gustavo")
    driver.find_element(By.NAME, "f4").send_keys("12032006")

    # Clica no botão de login diretamente
    login_button = driver.find_element(By.ID, "5")
    
    # Usando JavaScript para clicar no botão
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)

    # Preenche os campos de Unidade e Opção
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f3").send_keys("455")
        
    time.sleep(7)  # Atraso para carregar a nova aba
        
    abas = driver.window_handles  # Lista o número de abas abertas.

    # Muda o foco para a última aba (a nova aba)
    driver.switch_to.window(abas[-1])

    # Cria uma instância de ActionChains
    actions = ActionChains(driver)

    data_atual = datetime.now().strftime('%d%m%y')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f12")))
    input_element = driver.find_element(By.ID, "12")
    driver.execute_script("arguments[0].value = '';", input_element)
    time.sleep(0.3)
    driver.find_element(By.ID, "12").send_keys(data_atual)
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").send_keys("t")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").send_keys("e")
    time.sleep(0.3)
    input_element = driver.find_element(By.NAME, "f37")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.NAME, "f37").send_keys("g")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f38").send_keys("h")

    login_button = driver.find_element(By.ID, "40")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(0.2)
    actions.send_keys("1").perform()

    time.sleep(5)
    abas = driver.window_handles  # Lista o número de abas abertas.

    # Muda o foco para a última aba (a nova aba)
    driver.switch_to.window(abas[-1])
    time.sleep(2)

    # Captura o valor do seq da segunda linha da primeira coluna da tabela com o ID 'tblsr'
    tabela = driver.find_element(By.ID, "tblsr")
    linhas = tabela.find_elements(By.TAG_NAME, "tr")

    # Verifica se há pelo menos duas linhas (cabeçalho + dados)
    if len(linhas) > 1:
        seq_da_requisicao = linhas[1].find_element(By.TAG_NAME, "td").text  # Captura o valor da primeira coluna da segunda linha
        print(f"Seq da requisição: {seq_da_requisicao}")
    else:
        print("Não há linhas suficientes na tabela para capturar o seq.")

    # Clica no botão de atualização
    update_button = driver.find_element(By.ID, "2")
    time.sleep(60)
    driver.execute_script("arguments[0].click();", update_button)
    time.sleep(5)

    # Após a atualização, busca novamente a linha correspondente ao seq
    relatorios_atualizados = driver.find_elements(By.CSS_SELECTOR, "table#tblsr tr")
    relatorio_encontrado = None

    for relatorio in relatorios_atualizados[1:]:  # Ignora o cabeçalho da tabela
        seq_atual = relatorio.find_element(By.TAG_NAME, "td").text  # Captura o seq da primeira coluna
        if seq_atual == seq_da_requisicao:
            relatorio_encontrado = relatorio
            break

    # Se um relatório correspondente foi encontrado, clique no link (substituído por <u>) na mesma linha
    if relatorio_encontrado:
        link = relatorio_encontrado.find_element(By.TAG_NAME, "u")  # Supondo que o link esteja dentro de uma tag <u>
        driver.execute_script("arguments[0].click();", link)
        print("Clicou no link da requisição correspondente.")
    else:
        print("Nenhum relatório encontrado após a atualização.")

    time.sleep(10)  # Atraso para garantir que o download seja iniciado

finally:
    # Fecha o navegador
    driver.quit()