import time

from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from pywinauto.application import Application
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import ControleTerminal3270
import keyboard as kb
#from selenium.webdriver.chrome.options import Options


#options = Options()
#options.add_experimental_option("detach", True)
servico = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service=servico)
#
#
#Abrindo o Chrome
url = 'https://sei.economia.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI'
browser.get(url)
browser.maximize_window()
time.sleep(5)

#Realizando login
browser.find_element(By.ID, 'txtUsuario').send_keys('eduardo.marques@economia.gov.br')
browser.find_element(By.ID, 'pwdSenha').send_keys('Duduym985623ae.')
browser.find_element(By.XPATH, '//*[@id="selOrgao"]').send_keys('ME')
browser.find_element(By.ID, 'Acessar').click()
time.sleep(5)


# Gerenciar Marcador
# browser.switch_to.default_content()
# browser.switch_to.frame('ifrVisualizacao')
gerenciar_marcador = browser.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[6]/td[1]/a[2]').click()
time.sleep(3)

#Adicionar Marcador
browser.switch_to.default_content()
browser.switch_to.frame('ifrVisualizacao')
adicionar = browser.find_element(By.XPATH,'//*[@id="btnAdicionar"]').click()

#Marcador PROCESSO INSTRUÍDO
browser.switch_to.default_content()
browser.switch_to.frame('ifrVisualizacao')
marcador = browser.find_element(By.CLASS_NAME,'dd-selected').click()
processo_instruido = browser.find_element(By.XPATH,'//li[contains(text(), "Processo INSTRUÍDO")]').click()
salvar = browser.find_element(By.XPATH,'//*[@id="sbmSalvar"]').click()