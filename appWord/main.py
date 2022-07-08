import os
import re
import sys
import time
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout
from Formatter import Formatter
from Settings import Settings

import win32com.client as win32
from glob import glob

# Подкласс QMainWindow для настройки главного окна приложения
class MainWindow(QWidget):
    settings = {'ChangeNumber': True, 'ChangeDate': True, 'ChangeKavich': True, 'ChangeTN': True, 'ChangeTire': True
        , 'ChangePadeg': True, 'ChangeRF': True}

    def __init__(self):
        super().__init__()

        self.setWindowTitle("My App")

        self.btnOpenFile = QPushButton("Открыть файл")
        self.btnOpenFile.setCheckable(True)
        self.btnOpenFile.clicked.connect(self.the_button_was_clicked)

        self.btnFormat = QPushButton("Форматировать")
        self.btnFormat.setCheckable(True)

        self.btnSettings = QPushButton("Настройки")
        self.btnSettings.setCheckable(True)
        self.btnSettings.clicked.connect(self.open_settings)

        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.btnOpenFile, 0, 0)
        self.grid.addWidget(self.btnFormat, 1, 0)
        self.grid.addWidget(self.btnSettings, 2, 0)

        self.setLayout(self.grid)

        self.setGeometry(app.screens()[0].size().width() - 120, (app.screens()[0].size().height() // 2) - 150, 120, 300)

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

        self.show()

    def save_as_docx(self, path):

        # Открываем Word
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(path)
            doc.Activate()

            # Меняем расширение на .docx и добавляем в путь папку
            # для складывания конвертированных файлов
            # new_file_abs = str(os.path.abspath(path)).split("\\")
            # new_dir_abs = f"{new_file_abs[0]}\\{new_file_abs[1]}"
            # new_file_abs = f"{new_file_abs[0]}\\{new_file_abs[1]}\\doc_convert\\{new_file_abs[2]}"
            # new_file_abs = os.path.abspath(new_file_abs)
            # if not os.path.isdir(f'{new_dir_abs}\\doc_convert'):
            #     os.mkdir(f'{new_dir_abs}\\doc_convert')
            new_file_abs = re.sub(r'\.\w+$', '.docx', path)
            print(new_file_abs)
            # Сохраняем и закрываем
            word.ActiveDocument.SaveAs(new_file_abs, FileFormat=win32.constants.wdFormatXMLDocument)
            doc.Close(False)
        except:
            return str(path).split("\\")[-1]

    def get_word(self, flag):
        if flag:
            paths = sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.doc'))
        elif not flag:
            paths = sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.docx'))
        local_time = time.ctime(time.time()).split(' ')
        local_time = [i for i in local_time if i != '']
        local_time.pop(3)
        files = sorted(paths, key=os.path.getctime, reverse=True)
        new_files = []
        for f in files:
            print(f, "hui")
            time1 = time.ctime(os.path.getctime(f)).split(' ')
            time1 = [i for i in time1 if i != '']
            time1.pop(3)
            if time1 == local_time and (not '~$' in f.name):
                new_files.append(f)
        return new_files

    def the_button_was_clicked(self):
        # top = tkinter.Tk()
        # top.withdraw()
        # file_name = fd.askopenfilename(parent=top, filetypes=(("Word", "*.doc"), ("All files", "*.*")))
        # top.destroy()

        file_name = self.get_word(True)[0]
        print(file_name)
        self.save_as_docx(os.path.normpath(file_name))
        file_name = self.get_word(False)[0]
        print(file_name)
        #os.system('taskkill /f /im WINWORD.EXE') # закрывает все ВОРДЫ
        try:
            frm = Formatter(file_name, settings=self.settings)
            frm.Redact()
            #os.startfile(file_name) #открывает док
        except FileNotFoundError:
            print(file_name)
            return
        except Exception as exc:
            print(exc)

    def open_settings(self):
        self.settings_window = Settings(self.settings)
        self.settings_window.show()


# Мониторинг новый запущенных программ

app = QApplication(sys.argv)
window = MainWindow()

app.exec()