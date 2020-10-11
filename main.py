from kivy.app import App
from kivy.lang.builder import Builder
from kivy.properties import ObjectProperty
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.core.window import Window
from Convert_files import *
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.dropdown import DropDown

filelocal_ = None

class Gerenciador(ScreenManager):  # Gerencia a troca de telas
    pass


def voltar(window, key, *args):  # Retorna para o menu principal ao apertar esc, ou o botão voltar do android
    if key == 27:
        App.get_running_app().root.current = 'screen'
        return True

def CustonDropDown(BoxLayout):
    pass


class GerarRelatorio(Screen):
    def on_pre_enter(self):
        Window.bind(on_keyboard=voltar)  # Define a tecla esc como a função voltar, para retornar a tela anterior

    def on_pre_leave(self):
        Window.unbind(on_keyboard=voltar)  # Retira a função esc quando está no menu principal, assim quando for pressionado, o aplicativo fecha


class ConverterArquivos(Screen): 

    def on_pre_enter(self):
        Window.bind(on_keyboard=voltar)

    def on_pre_leave(self):
        Window.unbind(on_keyboard=voltar)

    def relatorio(self):
        global filelocal_
        arquivos = self.ids.Files.text
        filelocal_ = filelocal(arquivos)

    def convert(self):
        global filelocal_
        values = [self.ids.Files.text, self.ids.opentype.text, self.ids.closetype.text]
        try:
            Convert(filelocal_, values)
        except AttributeError:
            print('Selecione o Diretório')

class Bitworks(App):  # Herda da classe App de kivy.app para gerar o programa
    def build(self):
        Builder.load_string(open("widgets.kv", encoding="utf-8").read(), rulesonly=True)  # Abre o arquivo widgets.kv onde estará a interface do usuário e utiliza o encoding utf-8 para ler acentos
        return Gerenciador()

if __name__ == '__main__':
    Bitworks().run()
