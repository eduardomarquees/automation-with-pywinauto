import csv
import time
import glob
import os
from tkinter import *
import keyboard as kb
import ControleTerminal3270
from pywinauto import *
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
import pandas as pd


#Tempo UNIVERSAL
intervalo = 0.5


#Interagindo com SIAPENET
app = Application().connect(title_re="^Terminal 3270.*")
dlg = app.window(title_re="^Terminal 3270.*")
Acesso = ControleTerminal3270.Janela3270()
time.sleep(intervalo)
#dlg.type_keys('{F3}')
time.sleep(intervalo)
dlg.type_keys('{F2}')


#                    #EXTRAINDO VÍNCULOS...
dlg.type_keys('>CDCONVINC')
kb.press("ENTER")
lista_dados = []
dados_invalidos = []


nome_arquivo = "dados.csv"
nao_encontrado = "cpf_nao_encontrado.csv"
cpf_invalido = "cpfs_invalidos.csv"


#Abrindo arquivos para leitura de CPFS
data = pd.read_csv(nome_arquivo, dtype={'CPF': str})
for elemento_desejado in data['CPF']:
    print(elemento_desejado)
    time.sleep(intervalo)
    dlg.type_keys('{TAB}')
    dlg.type_keys(elemento_desejado)
    kb.press("ENTER")

    cpf_invalido = Acesso.pega_texto_siape(Acesso.copia_tela(), 24, 9, 24, 80).strip()
    if cpf_invalido == "CPF INVALIDO":
        dados_invalidos.append(elemento_desejado)
        dlg.type_keys('{TAB}')
        continue


    time.sleep(intervalo)

    erro_cpf = Acesso.pega_texto_siape(Acesso.copia_tela(), 24, 9, 24, 80).strip()
    if erro_cpf == "NAO EXISTEM DADOS PARA ESTA CONSULTA":
        lista_dados.append(elemento_desejado)
        dlg.type_keys('{TAB}')
        continue


    time.sleep(intervalo)
    kb.press("ENTER")
    # time.sleep(5)
    dlg.type_keys('x')
    kb.press("ENTER")
    time.sleep(intervalo)

    # Copiando nome é imprimindo o Vínculo
    encontrar_nome = Acesso.pega_texto_siape(Acesso.copia_tela(), 6, 12, 6, 80).strip()
    time.sleep(intervalo)
    dlg.type_keys('^p')
    time.sleep(intervalo)

    # Obtendo Janela de impressão
    app = Application().connect(title_re="Imprimir")
    time.sleep(intervalo)
    window = app.window(title_re="Imprimir")
    window.set_focus()
    kb.press("Enter")
    time.sleep(intervalo)

    # Obtendo Janela de salvar saída de impressão
    app = Application().connect(title_re='Salvar Saída de Impressão como')
    dlg = app[u'Salvar Saída de Impressão como']
    time.sleep(intervalo)
    window = app.window(title_re='Salvar Saída de Impressão como')
    window.set_focus()
    encontrar_nome = encontrar_nome.replace(" ", "{SPACE}")
    time.sleep(intervalo)
    dlg.type_keys('Vínculo' + "{SPACE}" + encontrar_nome)
    kb.press("Enter")
    time.sleep(intervalo)

    # Voltando para página de vínculo
    app = Application().connect(title_re="^Terminal 3270.*")
    dlg = app.window(title_re="^Terminal 3270.*")
    Acesso = ControleTerminal3270.Janela3270()
    dlg.type_keys('{F12}')
    #dlg.type_keys('{TAB}')


df = pd.DataFrame(lista_dados, columns=['CPF'])
df.to_csv('cpf_nao_encontrado.csv', index=False)

df2 = pd.DataFrame(dados_invalidos, columns=['CPFS Invalidos'])
df2.to_csv('cpfs_invalidos.csv', index=False)



