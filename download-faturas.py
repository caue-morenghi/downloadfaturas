from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time
from pathlib import Path
import zipfile
import shutil
import pandas as pd
import os

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)


# login
def login():
    driver.get("https://empresa.odontoprev.com.br/dashboard")
    time.sleep(2)
    driver.refresh()
    time.sleep(3)

    acessar_login_button = driver.find_element(
        By.XPATH, '/html/body/div[1]/div/div/div/header/div[2]/div/div[2]/div/div[2]').click()

    time.sleep(10)

    login_input = driver.find_element(By.XPATH, '//*[@id="input-74"]')
    login_input.click()
    login_input.send_keys("10444991")

    password_input = driver.find_element(By.XPATH, '//*[@id="input-77"]')
    password_input.click()
    password_input.send_keys("@B3lh1nh4")
    password_input.send_keys(Keys.RETURN)

    time.sleep(5)
    driver.switch_to.window(driver.window_handles[0])


# função de entrar na parte de financeiro do site e escolher a operadora ---> se a operadora atual já estiver selecionada, o código não faz nada, caso contrário, ele seleciona a próxima operadora e entra na tela "financeiro" novamente
def escolher_operadora(i_operadora):
    driver.get("https://empresa.odontoprev.com.br/dashboard/financeiro")
    time.sleep(5)
    driver.find_element(
        By.XPATH, '/html/body/div[1]/div/div/div[3]/div/div/button/span').click()
    time.sleep(3)
    driver.find_element(
        By.XPATH, '/html/body/div[1]/div/div/div[1]/header/div[2]/div/div[2]/div/div[1]/div/div[2]/div/div[1]').click()
    time.sleep(3)
    driver.find_element(
        By.XPATH, f'//div[@id="list-item-186-{i_operadora}"]').click
    time.sleep(5)

    try:
        driver.find_element(
            By.XPATH, '/html/body/div[1]/div/div/div[5]/div/div/div[1]/h2')
        driver.find_element(
            By.XPATH, '/html/body/div[1]/div/div/div[5]/div/div/div[2]/div[2]').click()
        time.sleep(15)
        driver.get("https://empresa.odontoprev.com.br/dashboard/financeiro")
        time.sleep(5)
        driver.find_element(
            By.XPATH, '/html/body/div[1]/div/div/div[3]/div/div/button/span').click()
    except NoSuchElementException:
        print("atual operadora já selecionada")


# processo de baixar os relatórios, extrair o arquivo zip e mover para a pasta de relatórios, extrair os arquivos zip internos e deletar os arquivos zip, e por fim, extrair os arquivos csv, ler e printar o conteúdo ---> está printando o conteúdo do csv, mas ainda não está fazendo nada com ele ---> deve-se pegar cada campo de cada linha do csv e inserir na tabela "linhas_faturas" do banco de dados
for i_operadora in range(1, 6):
    print(f"LOOP NÚMERO {i_operadora}")

    if i_operadora == 1:
        nome_operadora = 'MANSERV INVES'
    elif i_operadora == 2:
        nome_operadora = 'E-SERVICOS S'
    elif i_operadora == 3:
        nome_operadora == 'MANSERV-FACIL'
    elif i_operadora == 4:
        nome_operadora = 'MANSERV MANUT'
    elif i_operadora == 5:
        nome_operadora = 'LSI - LOGISTI'
        
    pasta_base_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    pasta_base_mesa = os.path.join(os.path.expanduser("~"), "Desktop")

    login()

    escolher_operadora(i_operadora=i_operadora)
    time.sleep(3)

    

    competencias = driver.find_elements(
        By.XPATH, '//*[@id="app"]/div[1]/main/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div/div/div[1]/div[1]/div/table/tbody/tr/td[1]')

    ano_atual_str = time.strftime("%Y", time.localtime())
    mes_atual_str = time.strftime("%m", time.localtime())
    mes_atual_int = int(time.strftime("%m", time.localtime()))
    rodagem = 0
    for competencia in competencias:
        rodagem += 1
        print(rodagem)
        data_competencia = competencia.text
        mes_competencia = int(data_competencia.split("/")[1])
        print(mes_atual_str, mes_competencia, ano_atual_str)
        if mes_competencia == mes_atual_int - 1:
            relatorios_btn_div = driver.find_element(
                By.XPATH, f'//*[@id="app"]/div[1]/main/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div/div/div[1]/div[1]/div/table/tbody/tr[{rodagem}]/td[7]/div/div[1]').click()
            time.sleep(4)
            selecionar_todos_btn = driver.find_element(
                By.XPATH, '//div[@class="v-input--selection-controls__ripple"][1]').click()
            time.sleep(2)
            baixar_btn = driver.find_element(
                By.XPATH, '//*[@id="app"]/div[5]/div/div/div/div[2]/div/button').click()
            time.sleep(10)

            # extrai arquivos e joga pra pasta da operadora
            # aqui, ele está movendo para uma pasta que está na area de trabalho, que tem dentro dela 5 pastas com os nomes das 5 operadoras (MANSERV INVES, E-SERVICOS S, MANSERV-FACIL, MANSERV MANUT, LSI - LOGISTI)
            arq_relatorio = Path(
                f"{pasta_base_downloads}\relatorios-selecionados.zip")
            if arq_relatorio.exists():
                print("arquivo encontrado")
                with zipfile.ZipFile(arq_relatorio, 'r') as zip_ref:
                    zip_ref.extractall(
                        f"{pasta_base_downloads}\relatorios-selecionados")
                    pastas_relatorios_selecionados = Path(
                        f"{pasta_base_downloads}\relatorios-selecionados")
                    pasta_relatorios_destino = Path(r"{}\relatorios".format(pasta_base_mesa))

                    if not pasta_relatorios_destino.exists():
                        os.makedirs(pasta_relatorios_destino)

                    if pastas_relatorios_selecionados.exists():
                        shutil.move(pastas_relatorios_selecionados,
                                    r"{}\relatorios\{}\{}\{}".format(pasta_base_mesa, ano_atual_str, mes_atual_str, nome_operadora))
                        print("arquivo movido")
                        print(mes_atual_str, ano_atual_str, nome_operadora, pasta_base_mesa, pasta_base_downloads)

                    # extrai arquivos zip internos (DIRF.zip ---> nesse contem o csv) e deleta os arquivos zip
                    pasta_final_relatorios_selecionados = Path(
                        r"{}\relatorios\{}\{}\{}\relatorios-selecionados".format(pasta_base_mesa, ano_atual_str, mes_atual_str, nome_operadora))
                    if pasta_final_relatorios_selecionados.exists():
                        arquivos_interiores = [arquivo for arquivo in pasta_final_relatorios_selecionados.iterdir(
                        ) if arquivo.is_file() and arquivo.suffix == '.zip']

                    for arq in arquivos_interiores:
                        with zipfile.ZipFile(arq, 'r') as zip_ref:
                            zip_ref.extractall(
                                pasta_final_relatorios_selecionados)
                        arq.unlink()

                # PARTE DA TABELA FATURAS (ainda tem que ser feito aqui)
                    # ----------
                # PARTE DA TABELA LINHA_TABELAS (deve ser concluída aqui)
                    # aqui, está sendo lido o csv e printado o conteúdo, mas deve-se pegar cada campo de cada linha e inserir na tabela "linhas_faturas" do banco de dados
                # csv_faturas = [arq for arq in pasta_final_relatorios_selecionados.iterdir(
                # ) if arq.is_file() and arq.suffix == '.csv']

                # for csv_arq_fatura in csv_faturas:
                #     df_faturas = pd.read_csv(
                #         csv_arq_fatura, encoding='ISO-8859-1', decimal=',', sep=';', dtype=str)
                #     print(df_faturas)

                arq_relatorio.unlink()

                time.sleep(2)
                # desmarca o input "Selecionar todos"
                try:
                    selecionar_todos_btn = driver.find_element(
                    By.XPATH, '//div[@class="v-input--selection-controls__ripple flue--text"]').click()

                    time.sleep(2)

                # Clica no "X" para fechar a janela de seleção de relatórios, AQUI ESTÁ DANDO ERRO: ELEMENTO NÃO INTERAGÍVEL
                #driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/div/div/button/span/i').click()
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[5]/div/div/span').click()
                except NoSuchElementException:
                    print("erro ao desmarcar o input 'Selecionar todos'")
            else:
                print("arquivo não encontrado")

        else:
            pass