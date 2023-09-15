from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import random
from selenium.common.exceptions import NoSuchElementException, TimeoutException


class WebAutomation:
    def __init__(self, browser):
        self.browser = browser

    def clica_elemento_by_xpath(self, iframe, path):
        """
        Clica em um elemento localizado por XPath dentro de um iframe.
        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.
        Returns:
            None
        """
        try:
            self.browser.switch_to.default_content()
            self.browser.switch_to.frame(iframe)
            self.browser.find_element(By.XPATH, path).click()
            self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao clicar no elemento: {e}")

    def inserir_texto_enter_by_id_in_iframe(self, iframe, element_id, text):
        """
        Insere um texto em um elemento localizado por ID dentro de um iframe e pressiona a tecla ENTER.
        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            element_id (str): ID do elemento.
            text (str): Texto a ser inserido no elemento.
        Returns:
            None
        """
        try:
            self.browser.switch_to.frame(iframe)
            wait = WebDriverWait(self.browser, 10)
            element = wait.until(EC.presence_of_element_located((By.ID, element_id)))
            element.clear()
            element.send_keys(text)
            element.send_keys(Keys.ENTER)
            self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao inserir texto: {e}")

    def acessar_site_chrome(self, url, tempo_espera):
        """
        Acessa um site usando o browser Chrome, maximiza a janela e aguarda pelo tempo especificado.

        Parâmetros:
        - url (str): Endereço do site que será acessado.
        - tempo_espera (int): Tempo em segundos que o método deve aguardar após acessar o site.

        Retorna:
        - Nenhum.
        """
        try:
            self.browser.get(url)
            self.browser.maximize_window()
            self.contagem_regressiva(tempo_espera)
        except Exception as e:
            print(f"Erro ao acessar o site {url}: {e}")

    def contagem_regressiva(self, segundos):
        """
        Realiza uma contagem regressiva de segundos exibindo mensagens.

        Args:
            segundos (int): O número de segundos para a contagem regressiva.

        Returns:
            None
        """
        try:
            for i in range(segundos, -1, -1):
                if i == 0:
                    print("Iniciando!")
                else:
                    print(f"Iniciando em {i} segundos...")
                sleep(1)
        except Exception as e:
            print(f"Erro durante a contagem regressiva: {e}")

    def open_browser(url, browser_type="firefox"):
        try:
            if browser_type == "firefox":
                driver = webdriver.Firefox()
            elif browser_type == "chrome":
                driver = webdriver.Chrome()
            else:
                print(f"Tipo de navegador {browser_type} não suportado!")
                return None

            driver.get(url)
            return driver
        except Exception as e:
            print(f"Erro ao abrir o browser: {e}")
            return None

    def tempo_aleatorio(self, inicio, fim):
        """
        Retorna um número aleatório entre os valores 'inicio' e 'fim'.

        Parâmetros:
        - inicio (int): Valor mínimo para a geração do número aleatório.
        - fim (int): Valor máximo para a geração do número aleatório.

        Retorna:
        - int: Número aleatório entre 'inicio' e 'fim'.
        """
        try:
            tempo = random.randint(inicio, fim)
            return tempo
        except Exception as e:
            print(f"Erro ao gerar tempo aleatório: {e}")
            return (inicio + fim) // 2  # retorna um valor médio entre 'inicio' e 'fim' em caso de erro

    def tela_aviso(self, path):
        """
        Fecha a janela de aviso se encontrada.
        Parâmetros:
        - path (str): XPath do elemento a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.implicitly_wait(30)
            if self.browser.find_element(By.XPATH, path):
                self.browser.find_element(By.XPATH, path).click()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao fechar janela de aviso: {e}")

    def sel_unidade(self, path):
        """
        Seleciona uma unidade no browser.
        Parâmetros:
        - path (str): XPath do elemento a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.implicitly_wait(30)
            self.browser.find_element(By.XPATH, path).click()
            sleep(self.tempo_aleatorio())
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao selecionar unidade: {e}")

    def visua_detal(self, path):
        """
        Seleciona a visualização detalhada no controle de processos.
        Parâmetros:
        - path (str): XPath do elemento a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            sleep(self.tempo_aleatorio())
            self.browser.find_element(By.XPATH, path).click()
            sleep(self.tempo_aleatorio())
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao selecionar visualização detalhada: {e}")

    def procura_marcador(self, path):
        """
        Seleciona a visualização por marcadores e clica no marcador especificado.
        Parâmetros:
        - path (str): XPath do marcador a ser clicado.
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.implicitly_wait(30)
            self.browser.find_element(By.XPATH, path).click()
            sleep(self.tempo_aleatorio())
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao procurar marcador: {e}")

    def procura_proc_esp(self, path, processo):
        """
        Busca por um processo específico usando o XPath fornecido.
        Parâmetros:
        - path (str): XPath do campo de pesquisa do processo.
        - processo (str): Número ou identificação do processo a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            tempo = random.randint(3, 7)
            self.browser.implicitly_wait(30)
            self.browser.find_element(By.XPATH, path).click()
            busca_proc = self.browser.find_element(By.XPATH, path)
            busca_proc.send_keys(processo)
            sleep(tempo)
            busca_proc.send_keys(Keys.ENTER)
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao procurar processo específico: {e}")
