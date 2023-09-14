import time
import glob
import os
from tkinter import *
import keyboard as kb
import ControleTerminal3270
from pywinauto import *
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
intervalo = 1.5
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
nome_arquivo = "cpfs.txt"
cpfs = []
with open(nome_arquivo, "r") as arquivo:
    cpfs = arquivo.read().splitlines()
for cpf in cpfs:
    #print("CPF:", cpf)
    time.sleep(intervalo)
    dlg.type_keys('{TAB}')
    dlg.type_keys(cpf)
    kb.press("ENTER")
    time.sleep(intervalo)
    erro_cpf = Acesso.pega_texto_siape(Acesso.copia_tela(), 24, 2, 24, 80).strip()
    print(erro_cpf)
    if erro_cpf == "(4057) NAO EXISTEM DADOS PARA ESTA CONSULTA":
        print(cpf)
        dlg.type_keys('{TAB}')
        continue
    time.sleep(3)
    kb.press("ENTER")
    #time.sleep(5)
    dlg.type_keys('x')
    kb.press("ENTER")
    time.sleep(intervalo)
    encontrar_nome = Acesso.pega_texto_siape(Acesso.copia_tela(),6,12,6,80).strip()
    #print(encontrar_nome)
    time.sleep(intervalo)
    dlg.type_keys('^p')
    time.sleep(intervalo)

    #Obtendo Janela
    app = Application().connect(title_re="Imprimir")
    time.sleep(intervalo)
    window = app.window(title_re="Imprimir")
    window.set_focus()
    kb.press("Enter")
    time.sleep(intervalo)

    #Obtendo Janela
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

    #Voltando para página de vínculo
    app = Application().connect(title_re="^Terminal 3270.*")
    dlg = app.window(title_re="^Terminal 3270.*")
    Acesso = ControleTerminal3270.Janela3270()
    dlg.type_keys('{F12}')













