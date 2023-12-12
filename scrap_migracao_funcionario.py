from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import csv
import time

# Caminho para salvar o arquivo CSV
caminho_csv = r"Caminho de onde deseja salvar o arquivo\funcionarios.csv"


#############################################################################
#                    PARTE DO CÓDIGO PARA SER ALTERADA                      #
#############################################################################
field_senha = "senha do cliente"
field_unidade_negocio = "unidade do cliente"
# 1 para versão mais básica do sistema, 2 para a versão completa 
system_version = "tipo do sistema"
#############################################################################


# Clicar no botão conforme a opção
if system_version == "1":

    # Configuração do navegador
    driver = webdriver.Chrome()
    driver.get("https://app.tecnoponto.com/login.jsf") 

    # Preencher login
    user_field = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[2]/div[1]/input")
    user_field.send_keys("user")

    # Preencher senha
    password_field = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[2]/div[2]/input").send_keys(field_senha)

    # Preencher Unidade de Negócio
    company_field = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[2]/div[3]/input").send_keys(field_unidade_negocio)

    # Clicar no botão de login
    login_button = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[5]/div/div/div/button/span[2]")
    login_button.click()

    # Aguardar a mudança de URL após o login
    driver.implicitly_wait(20)

    # Atualizar a url do site
    account_url = driver.current_url

    #Clicar no campo de cadastro
    register_option = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[2]/ul/form/div/div[1]/h3/a")
    register_option.click()

    #Clicar no campo de funcionário
    employee_option = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[2]/ul/form/div/div[1]/div/ul/li[6]/a/span")
    employee_option.click()

    # Aguardar a mudança de URL após o login
    driver.implicitly_wait(20)

    # Atualizar a url do site
    account_url = driver.current_url

    # Clicar no botão consultar
    consult_option = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[1]/td/table/tbody/tr[1]/td[1]/button/span[2]")
    consult_option.click()

    time.sleep(10)

    # Localizar o elemento select
    select_element = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[4]/td/table/tbody/tr/td/div/div[2]/select[2]")

    # Criar um objeto Select
    select = Select(select_element)

    # Selecionar a opção desejada pelo índice (índice 2 para a terceira opção, já que os índices começam em 0)
    select.select_by_index(2)

    time.sleep(10)

    #Puxa os dados
    xpath_index = 1

    cabecalho = ["NOME", "CPF", "PIS", "RG", "MATRÍCULA", "DATA ADMISSÃO", "DATA DEMISSÃO", "INÍCIO MARCAÇÃO", "DOCUMENTO EMPRESA"]
    with open(caminho_csv, "a", newline="", encoding="utf-8") as arquivo_csv:
            escritor_csv = csv.writer(arquivo_csv)
            escritor_csv.writerow(cabecalho)

    # Loop para clicar em todos os botões na tabela
    while True:
        edit_button_xpath = f"/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[4]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[{xpath_index}]/td[2]/button/span[1]"

        try:
            # Clicar no botão usando o XPath
            edit_button = driver.find_element(By.XPATH, edit_button_xpath)
            edit_button.click()

            # Lista de XPaths dos elementos
            xpaths = [
            
            #Nome
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[1]/div/div[2]/div/input",
            #CPF
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div[2]/div/input",
            #PIS
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[2]/div[1]/div/div/div[2]/div/input",
            #RG
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[1]/div[2]/div/div/div[2]/div/input",
            #Matricula
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[1]/div[1]/div/div/div[2]/div/input",
            #Data admissão
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[5]/div/div[1]/div[1]/div/div/div/div[1]/div/div/div[2]/div/span/input",
            #Data demissão
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[5]/div/div[1]/div[1]/div/div/div/div[2]/div/div/div[2]/div/span/input",
            #Data inicio marcação
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[5]/div/div[2]/div[1]/div/div/div/div[1]/div/div/div[2]/div/span/input",
            #Empresa
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[3]/div[1]/div/div/div[2]/div/span/input"
    ]
            # Lista para armazenar os dados
            dados = []

            # Iterar sobre os XPaths e adicionar os valores à lista
            for xpath in xpaths:
                elemento = driver.find_element(By.XPATH, xpath)
                valor = elemento.get_attribute("value")
                dados.append(valor)

            # Salvar os dados em um arquivo CSV
            with open(caminho_csv, "a", newline="", encoding="utf-8") as arquivo_csv:
                    escritor_csv = csv.writer(arquivo_csv)
                    escritor_csv.writerow(dados)

            # Clicar no botão de retorno
            return_button = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/button/span[2]")
            return_button.click()

            # Incrementar o índice para o próximo botão
            xpath_index += 1

        except NoSuchElementException:

            # Se não houver mais botões, sair do loop
            break

    driver.quit()
    print("Dados de funcionarios salvo em arquivo CSV")

elif system_version == "2":
    driver.find_element(By.XPATH, button_xpath).click()
    # Configuração do navegador
    driver = webdriver.Chrome()
    driver.get("https://app.tecnoponto.com/login.jsf") 

    # Preencher login
    user_field = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[2]/div[1]/input")
    user_field.send_keys("user")

    # Preencher senha
    password_field = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[2]/div[2]/input").send_keys(field_senha)

    # Preencher Unidade de Negócio
    company_field = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[2]/div[3]/input").send_keys(field_unidade_negocio)

    # Clicar no botão de login
    login_button = driver.find_element(By.XPATH, "/html/body/div[1]/form/div/div/div[1]/div[1]/div[5]/div/div/div/button/span[2]")
    login_button.click()

    # Aguardar a mudança de URL após o login
    driver.implicitly_wait(20)

    # Atualizar a url do site
    account_url = driver.current_url

    #Clicar no campo de cadastro
    register_option = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[2]/ul/form/div/div[1]/h3/a")
    register_option.click()

    #Clicar no campo de funcionário
    employee_option = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[2]/ul/form/div/div[1]/div/ul/li[10]/a/span")
    employee_option.click()

    # Aguardar a mudança de URL após o login
    driver.implicitly_wait(20)

    # Atualizar a url do site
    account_url = driver.current_url

    # Clicar no botão consultar
    consult_option = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[1]/td/table/tbody/tr[1]/td[1]/button/span[2]")
    consult_option.click()

    time.sleep(10)

    # Localizar o elemento select
    select_element = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[4]/td/table/tbody/tr/td/div/div[2]/select[2]")

    # Criar um objeto Select
    select = Select(select_element)

    # Selecionar a opção desejada pelo índice (índice 2 para a terceira opção, já que os índices começam em 0)
    select.select_by_index(2)

    time.sleep(10)

    #Puxa os dados
    xpath_index = 1

    cabecalho = ["NOME", "CPF", "PIS", "RG", "MATRÍCULA", "DATA ADMISSÃO", "DATA DEMISSÃO", "INÍCIO MARCAÇÃO", "DOCUMENTO EMPRESA"]
    with open(caminho_csv, "a", newline="", encoding="utf-8") as arquivo_csv:
            escritor_csv = csv.writer(arquivo_csv)
            escritor_csv.writerow(cabecalho)

    # Loop para clicar em todos os botões na tabela
    while True:
        edit_button_xpath = f"/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[4]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[{xpath_index}]/td[2]/button/span[1]"

        try:
            # Clicar no botão usando o XPath
            edit_button = driver.find_element(By.XPATH, edit_button_xpath)
            edit_button.click()

            # Lista de XPaths dos elementos
            xpaths = [
            
            #Nome
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[1]/div/div[2]/div/input",
            #CPF
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[2]/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div[2]/div/input",
            #PIS
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[2]/div[1]/div/div/div[2]/div/input",
            #RG
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[1]/div[2]/div/div/div[2]/div/input",
            #Matricula
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[1]/div[1]/div/div/div[2]/div/input",
            #Data admissão
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[5]/div/div[1]/div[1]/div/div/div/div[1]/div/div/div[2]/div/span/input",
            #Data demissão
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[5]/div/div[1]/div[1]/div/div/div/div[2]/div/div/div[2]/div/span/input",
            #Data inicio marcação
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[5]/div/div[2]/div[1]/div/div/div/div[1]/div/div/div[2]/div/span/input",
            #Empresa
            "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[3]/td/div/div/div[3]/div/div/div/div[1]/div[3]/div/div[3]/div[1]/div/div/div[2]/div/span/input"
    ]
            # Lista para armazenar os dados
            dados = []

            # Iterar sobre os XPaths e adicionar os valores à lista
            for xpath in xpaths:
                elemento = driver.find_element(By.XPATH, xpath)
                valor = elemento.get_attribute("value")
                dados.append(valor)

            # Salvar os dados em um arquivo CSV
            with open(caminho_csv, "a", newline="", encoding="utf-8") as arquivo_csv:
                    escritor_csv = csv.writer(arquivo_csv)
                    escritor_csv.writerow(dados)

            # Clicar no botão de retorno
            return_button = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/form[16]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/button/span[2]")
            return_button.click()

            # Incrementar o índice para o próximo botão
            xpath_index += 1

        except NoSuchElementException:

            # Se não houver mais botões, sair do loop
            break

    driver.quit()
    print("Dados de funcionarios salvo em arquivo CSV")
else:
    print("Opção inválida")

