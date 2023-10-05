import time
import pandas as pd
import csv
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from ControleChrome import WebAutomation
from pywinauto.application import Application
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
servico = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service=servico)


#Abrindo o Chrome
url = 'https://sei.economia.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI'
browser.get(url)
browser.maximize_window()
time.sleep(5)

#Realizando login
browser.find_element(By.ID, 'txtUsuario').send_keys('eduardo.marques@economia.gov.br')
browser.find_element(By.ID, 'pwdSenha').send_keys(('985623ee..'))
browser.find_element(By.XPATH, '//*[@id="selOrgao"]').send_keys('ME')
browser.find_element(By.ID, 'Acessar').click()
time.sleep(10)

#Clicando em Visualização detalhada
browser.find_element(By.XPATH, '//*[@id="divFiltro"]/div[1]/a').click()
time.sleep(5)

#Clicando em ver por marcadores
browser.find_element(By.XPATH, '//*[@id="divFiltro"]/div[4]/a').click()
time.sleep(5)

#Selecionando Marcador
browser.find_element(By.XPATH, '//*[@onclick="filtrarMarcador(77951)"]').click()
time.sleep(10)

indice_atual = 1  # começa a partir do segundo item porque o índice 0 é o cabeçalho
lista_dados = []
primeira_pagina = True
while True:
    try:
        tabela_processos = browser.find_element(By.TAG_NAME, 'table')
        linhas = tabela_processos.find_elements(By.TAG_NAME, 'tr')

    except Exception as e:
        break

    if indice_atual >= len(linhas):
        try:
            next_button = browser.find_element(By.XPATH, '//*[@id="lnkInfraProximaPaginaSuperior"]/img')
            next_button.click()
            time.sleep(2)
            indice_atual = 1  # Reset do índice para a nova página
        except NoSuchElementException:
            break  # Sai do loop se não houver mais páginas

    linha = linhas[indice_atual]

    try:
        celulas = linha.find_elements(By.TAG_NAME, 'td')
        dados_linha = [celulas[2].text, celulas[4].text]
        # print(f"Capturando dados da linha {indice_atual}: {dados_linha}")  # Imprime os dados para teste
        lista_dados.append(dados_linha)  # Adiciona à lista


    except Exception as e:
        print("erro")

    indice_atual += 1

df = pd.DataFrame(lista_dados, columns=['Processo', 'CPF'])
df.to_csv('dados.csv', index=False)

controle_bot = Application









