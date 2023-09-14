from pywinauto.application import Application
from pywinauto import clipboard
import time


class Janela3270:

    def __init__(self):
        self.delay = 0.1

    def copia_tela(self):
        app = Application().connect(title_re="^Terminal 3270.*")
        dlg = app.top_window()
        dlg.type_keys('^a')
        time.sleep(self.delay)
        dlg.type_keys('^c')
        time.sleep(self.delay)
        tela = clipboard.GetData()
        return tela

    @staticmethod
    def pega_texto_siape(tela, L1, C1, L2, C2):
        # pegar um trecho do texto da tela que vai das coordenadas L1,C1 a L2,C2
        D1 = (L1 - 1) * 82 + C1 - 1
        D2 = (L2 - 1) * 82 + C2
        return tela[D1:D2]
