from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd


driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes")

try:
    driver.find_element(By.XPATH, "//button[contains(text(), 'Prosseguir')]").click()
except NoSuchElementException:
    pass

try:
    iframe = driver.find_element(By.CSS_SELECTOR, "iframe[src='https://ptax.bcb.gov.br/ptax_internet/consultaBoletim.do?method=exibeFormularioConsultaBoletim']")
    driver.switch_to.frame(iframe)
    botao_pesquisar = driver.find_element(By.CLASS_NAME, "botao")  
    botao_pesquisar.click()
    driver.switch_to.default_content()
except NoSuchElementException:
    pass

try:
    iframe = driver.find_element(By.CSS_SELECTOR, "iframe[src='https://ptax.bcb.gov.br/ptax_internet/consultaBoletim.do?method=exibeFormularioConsultaBoletim']")
    driver.switch_to.frame(iframe)

    tabela = driver.find_element(By.CLASS_NAME, "tabela")
    linhas = tabela.find_elements(By.TAG_NAME, "tr")

    dados_tabela = []
    for linha in linhas:
        colunas = linha.find_elements(By.TAG_NAME, "td")
        dados = [coluna.text for coluna in colunas]
        if dados:
            dados_tabela.append(dados)

    driver.switch_to.default_content()

    
    colunas = ["Data", "Tipo", "Compra", "Venda"]
    df = pd.DataFrame(dados_tabela, columns=colunas)
    df = df[:-1]
    print(df) 
    
except NoSuchElementException:
    print("Tabela ou iframe não encontrada.")

time.sleep(10)
driver.quit()
