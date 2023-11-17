import os
import time
from selenium.webdriver.support.ui import Select
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from pywinauto.application import Application
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
browser.find_element(By.ID, 'txtUsuario').send_keys('')
browser.find_element(By.ID, 'pwdSenha').send_keys('')
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

browser.implicitly_wait(3000)


df = pd.DataFrame(lista_dados, columns=['Processo', 'CPF'])
df.to_csv('dados.csv', index=False)



#ENTRANDO NO SIAPENET

#Tempo UNIVERSAL
intervalo = 1.3


#Interagindo com SIAPENET
app = Application().connect(title_re="^Terminal 3270.*")
dlg = app.window(title_re="^Terminal 3270.*")
Acesso = ControleTerminal3270.Janela3270()
time.sleep(intervalo)
dlg.type_keys('{F3}')
time.sleep(intervalo)
dlg.type_keys('{F2}')


                   #EXTRAINDO VÍNCULOS...
dlg.type_keys('>CDCONVINC')
kb.press("ENTER")
lista_dados = []
dados_invalidos = []


nome_arquivo = "dados.csv"
nao_encontrado = "cpf_nao_encontrado.csv"
cpf_invalido = "cpfs_invalidos.csv"


# #Abrindo arquivos para leitura de CPFS
data = pd.read_csv(nome_arquivo, dtype={'CPF': str})
cpfs = data['CPF'].tolist()
processos = data['Processo'].tolist()

for elemento_desejado, processos in zip(cpfs, processos):
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
        # dlg.type_keys('^p')
        #dlg.type_keys('{TAB}')

    else:
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

    # Obtém o diretório atual do script
    diretorio = os.path.abspath(os.path.dirname(__file__))

    # Cria uma subpasta chamada 'output' dentro do diretório atual
    pasta_saida = os.path.join(diretorio, 'output')
    os.makedirs(pasta_saida, exist_ok=True)

    # Obtém a lista de arquivos na subpasta 'output'
    lista_arquivos = [arq for arq in os.listdir(pasta_saida) if os.path.isfile(os.path.join(pasta_saida, arq))]


    processo = processos.replace(".", "").replace("/", "").replace("-", "")


    # Obtendo Janela de salvar saída de impressão
    app = Application().connect(title_re='Salvar Saída de Impressão como')
    dlg = app[u'Salvar Saída de Impressão como']
    time.sleep(intervalo)
    window = app.window(title_re='Salvar Saída de Impressão como')
    window.set_focus()
    encontrar_nome = encontrar_nome.replace(" ", "{SPACE}")
    time.sleep(intervalo)
    nomes = ('Vínculo' + "{SPACE}" + encontrar_nome + "{SPACE}" + processo)
    arquivo = os.path.join(pasta_saida, nomes)
    dlg.type_keys(arquivo)
    kb.press("Enter")
    time.sleep(intervalo)

    # Voltando para página de vínculo
    app = Application().connect(title_re="^Terminal 3270.*")
    dlg = app.window(title_re="^Terminal 3270.*")
    Acesso = ControleTerminal3270.Janela3270()
    dlg.type_keys('{F3}')
    dlg.type_keys('{F2}')
    dlg.type_keys('>CDCONVINC')
    kb.press('Enter')


#app.kill()
df = pd.DataFrame(lista_dados, columns=['CPF'])
df.to_csv('cpf_nao_encontrado.csv', index=False)

df2 = pd.DataFrame(dados_invalidos, columns=['CPFS Invalidos'])
df2.to_csv('cpfs_invalidos.csv', index=False)



data_atual = datetime.now().strftime("%d/%m/%Y")
conjuntos_caracters = []
# Obtém o diretório atual do script
diretorio = os.path.abspath(os.path.dirname(__file__))
# Concatena o diretório com a subpasta 'pdfs'
pasta = os.path.join(diretorio, 'output')
# Obtém a lista de arquivos na subpasta 'pdfs'
lista_arquivos = [arq for arq in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, arq))]
indice = 0

# Loop para encontrar e selecionar um arquivo PDF
for arquivo in lista_arquivos:
    if arquivo.endswith('.pdf'):
        conjuntos_caracters = arquivo.split('.pdf')[0][-17:]
        localizar_processo = browser.find_element(By.ID, 'txtPesquisaRapida')
        localizar_processo.send_keys(conjuntos_caracters)
        localizar_processo.send_keys(Keys.ENTER)
        time.sleep(10)

        # Clicando em incluir processo
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        browser.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[1]/img').click()
        time.sleep(5)

        # Procurando por Externo
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        procurar_externo = browser.find_element(By.XPATH, '//*[@id="tblSeries"]/tbody/tr[1]/td/a[2]')
        procurar_externo.click()

        #Tipo do Documento
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        documento = browser.find_element(By.XPATH,'//*[@id="selSerie"]/option[10]').click()

        #Data do Documento
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        data = browser.find_element(By.XPATH,'//*[@id="txtDataElaboracao"]').send_keys(data_atual)
        time.sleep(3)

        # Formato
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        formato = browser.find_element(By.XPATH, '//*[@id="divOptNato"]/div/label')
        formato.click()
        time.sleep(3)

        #Nome na Árvore
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        arvore = browser.find_element(By.ID,'txtNomeArvore')
        arvore.send_keys('CDCONVINC')

        # Localizando nível de acesso, clicando em restrito
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        drop = browser.find_element(By.XPATH, '//*[@id="lblRestrito"]')
        drop.click()
        time.sleep(3)

        # Hipótese Legal
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        browser.implicitly_wait(15)
        info_pessoal = Select(browser.find_element(By.XPATH, '//*[@id="selHipoteseLegal"]'))
        info_pessoal.select_by_visible_text("Informação Pessoal (Art. 31 da Lei nº 12.527/2011)")
        time.sleep(3)

        # Clicando em Anexar arquivo
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        clicando_anexo = browser.find_element(By.XPATH, '//*[@id="lblArquivo"]')
        clicando_anexo.click()
        time.sleep(3)

        # Abre a janela "Abrir"
        app = Application().connect(title_re="Abrir")
        time.sleep(0.5)
        dlg = app.window(title_re="Abrir")
        time.sleep(1.5)


        lista_arquivos[indice] = lista_arquivos[indice].replace(" ", "{SPACE}")
        caminho = os.path.join(pasta, lista_arquivos[indice])
        dlg.type_keys(caminho)
        dlg.type_keys('{ENTER}')
        time.sleep(5)
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        save = browser.find_element(By.XPATH,'//*[@id="btnSalvar"]')
        save.click()
        time.sleep(5)

        #Arvore de processo
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrArvore')
        clicar_processo = browser.find_element(By.CLASS_NAME,'infraArvoreNo').click()

        # Gerenciar Marcador
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        gerenciar_marcador = browser.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[22]/img').click()
        time.sleep(3)

        #Adicionar Marcador
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        adicionar = browser.find_element(By.XPATH,'//*[@id="btnAdicionar"]').click()

        #Marcador PROCESSO INSTRUÍDO
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        marcador = browser.find_element(By.CLASS_NAME,'dd-selected').click()
        processo_instruido = browser.find_element(By.XPATH,'//*[@id="selMarcador"]/ul/li[7]/a').click()
        salvar = browser.find_element(By.XPATH,'//*[@id="sbmSalvar"]').click()

        #Arvore Processo
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrArvore')
        clicar_processo = browser.find_element(By.CLASS_NAME, 'infraArvoreNoSelecionado').click()

        #Controle de processo
        browser.switch_to.default_content()
        browser.switch_to.frame('ifrVisualizacao')
        controle_processo = browser.find_element(By.XPATH,'//*[@id="divArvoreAcoes"]/a[24]/img').click()
        indice +=1














