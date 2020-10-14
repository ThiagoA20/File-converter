from kivy.app import App
from kivy.lang.builder import Builder
from kivy.properties import ObjectProperty
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.core.window import Window
from Convert_files import *
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.dropdown import DropDown

filelocal_ = None
savefile_ = None


"Manages the screen exchange"
class Manager(ScreenManager):
    pass


"returns to the main menu by pressing esc, or the android back button"
def back(window, key, *args):
    if key == 27:
        App.get_running_app().root.current = 'screen'
        return True

def CustonDropDown(BoxLayout):
    pass


class ConvertFiles(Screen): 
    def on_pre_enter(self):
        Window.bind(on_keyboard=back)

    def on_pre_leave(self):
        Window.unbind(on_keyboard=back)

    "Catch how much files user wants to convert, from widgets.kv, then pass as a parameter to the filelocal function in Convert_files.py"
    def report(self):
        global filelocal_
        arquivos = self.ids.Files.text  
        filelocal_ = filelocal(arquivos)

    "send the parameters received from the user interface to the Convert function of Convert_files.py and check filelocal_, if none returns a message asking for the location, else  performs the conversion then set filelocal_ to none for a new user operation"
    def convert(self):
        global filelocal_
        try:
            values = [self.ids.Files.text, self.ids.opentype.text, self.ids.closetype.text]
            Convert(filelocal_, values)
            filelocal_ = None
        except:
            if filelocal_ == None:
                print("Select the File local!")
            """
            elif savefile_ == None:
                print("Select the File local!")
            """


"inherits from the app class of kivy.app to generate the program"
class Bitworks(App):
    "opens the widgets.kv file where the user interface will be and uses utf-8 encoding to read accents"
    def build(self):
        Builder.load_string(open("widgets.kv", encoding="utf-8").read(), rulesonly=True)
        return Manager()

if __name__ == '__main__':
    Bitworks().run()
