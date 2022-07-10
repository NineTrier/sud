import os
import re
import sys
import time
import tkinter
from tkinter import filedialog as fd
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout
from PyQt5.uic.properties import QtGui, QtWidgets

from Formatter import Formatter
from Settings import Settings
import json
import win32com.client as win32
from glob import glob


# Подкласс QMainWindow для настройки главного окна приложения
class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.H, self.W = 150, 300
        self.settings = Settings()
        self.create_ui()

        self.show()

    def create_ui(self):
        print("Формирую окно...", end=" ")
        self.setWindowTitle("Redact Word")

        self.btnOpenFile = QPushButton("Откр. посл. документ")
        self.btnOpenFile.setCheckable(True)
        self.btnOpenFile.clicked.connect(self.the_button_was_clicked)

        self.btnOpenFileWithChoice = QPushButton("Выбрать документ")
        self.btnOpenFileWithChoice.setCheckable(True)
        self.btnOpenFileWithChoice.clicked.connect(self.the_button_was_clicked_with_choice)

        self.btnFormat = QPushButton("Форматировать")
        self.btnFormat.setCheckable(True)

        self.btnSettings = QPushButton("Настройки")
        self.btnSettings.setCheckable(True)
        self.btnSettings.clicked.connect(self.open_settings)

        self.setFixedSize(self.H, self.W)
        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.btnOpenFile, 0, 0)
        self.grid.addWidget(self.btnOpenFileWithChoice, 1, 0)
        self.grid.addWidget(self.btnFormat, 2, 0)
        self.grid.addWidget(self.btnSettings, 3, 0)

        self.setLayout(self.grid)

        self.setGeometry(app.screens()[0].size().width() - self.H,
                         (app.screens()[0].size().height() // 2) - (self.W // 2), self.H, self.W)

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        print("Готово!")

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
            # Сохраняем и закрываем
            word.ActiveDocument.SaveAs(new_file_abs, FileFormat=win32.constants.wdFormatXMLDocument)
            doc.Close(False)
            return True
        except:
            return False

    def get_word(self):
        paths = sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.doc'))
        paths += sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.docx'))
        print(paths)
        local_time = time.ctime(time.time()).split(' ')
        local_time = [i for i in local_time if i != '']
        local_time.pop(3)
        files = sorted(paths, key=os.path.getctime, reverse=True)
        new_files = []
        print(local_time)
        for f in files:
            time1 = time.ctime(os.path.getctime(f)).split(' ')
            time1 = [i for i in time1 if i != '']
            time1.pop(3)
            print(time1)
            if time1 == local_time and '~$' not in f.name:
                new_files.append(f)
        print(new_files)
        return new_files if len(new_files) > 0 else [""]

    def the_button_was_clicked_with_choice(self):
        top = tkinter.Tk()
        top.withdraw()
        file_name = fd.askopenfilename(parent=top,
                                       filetypes=(("Word", "*.doc"), ("Word", "*.docx"), ("All files", "*.*")))
        top.destroy()
        try:
            frm = Formatter(file_name, settings=self.settings)
            frm.Redact()
            # os.startfile(file_name) #открывает док
        except FileNotFoundError:
            print(file_name)
            return
        except Exception as exc:
            print(exc)

    def the_button_was_clicked(self):
        file_name = self.get_word()[0]
        print(file_name)
        if '.docx' not in file_name:
            print('Форматирую в .docx...')
            if self.save_as_docx(os.path.normpath(file_name)):
                print("Готово!")
            else:
                print("Не удалось реформатировать файл :(")
        # os.system('taskkill /f /im WINWORD.EXE') # закрывает все ВОРДЫ
        try:
            frm = Formatter(file_name, settings=self.settings)
            frm.Redact()
            # os.startfile(file_name) #открывает док
        except FileNotFoundError:
            return
        except Exception as exc:
            print(exc)

    def open_settings(self):
        self.settings.show()


# Мониторинг новый запущенных программ

app = QApplication(sys.argv)
window = MainWindow()

app.exec()
