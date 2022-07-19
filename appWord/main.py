import os
import re
import sys
import time
import win32com.client as win32
import socket

import tkinter
from tkinter import filedialog as fd

from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QApplication, QPushButton, QGridLayout
from PyQt5 import QtGui

from Formatter import Formatter
from Settings import Settings


# Подкласс QMainWindow для настройки главного окна приложения
class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.W, self.H = 75, 200
        self.WButton, self.HButton = 50, 45
        self.settings = Settings()
        self.create_ui()

        self.show()

    #Создание интерфейса
    def create_ui(self):
        print("Формирую окно...", end=" ")
        self.setWindowTitle("Форматтер")

        self.btnOpenFile = QPushButton("САД")
        self.btnOpenFile.setCheckable(True)
        self.btnOpenFile.clicked.connect(self.the_button_was_clicked)
        self.btnOpenFile.setFixedHeight(self.HButton)
        self.btnOpenFile.setFixedWidth(self.WButton)

        self.btnOpenFileWithChoice = QPushButton("Док.")
        self.btnOpenFileWithChoice.setCheckable(True)
        self.btnOpenFileWithChoice.clicked.connect(self.the_button_was_clicked_with_choice)
        self.btnOpenFileWithChoice.setFixedHeight(self.HButton)
        self.btnOpenFileWithChoice.setFixedWidth(self.WButton)

        self.btnSettings = QPushButton("Нас.")
        self.btnSettings.setCheckable(True)
        self.btnSettings.clicked.connect(self.open_settings)
        self.btnSettings.setFixedHeight(self.HButton)
        self.btnSettings.setFixedWidth(self.WButton)

        self.setFixedSize(self.W, self.H)
        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.btnOpenFile, 0, 0)
        self.grid.addWidget(self.btnOpenFileWithChoice, 1, 0)
        self.grid.addWidget(self.btnSettings, 2, 0)

        self.setLayout(self.grid)

        self.setGeometry(app.screens()[0].size().width() - self.W,
                         (app.screens()[0].size().height() // 2) - (self.H // 2), self.W, self.H)

        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.CustomizeWindowHint | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        print("Готово!")

    #Меняем разрешение с doc на docx и сохраняем
    def save_as_docx(self, path):
        # Открываем Word
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(path)
            doc.Activate()
            # Меняем расширение на .docx и добавляем в путь папку
            # для складывания конвертированных файлов
            new_file_abs = re.sub(r'\.\w+$', '.docx', path)
            # Сохраняем и закрываем
            word.ActiveDocument.SaveAs(new_file_abs, FileFormat=win32.constants.wdFormatXMLDocument)
            doc.Close(False)
            return True, new_file_abs
        except:
            return False, ""

    #Получаем последний открытый документ word c папки temp
    def get_word(self):
        paths = sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.doc'))
        paths += sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.docx'))
        local_time = time.ctime(time.time()).split(' ')
        local_time = [i for i in local_time if i != '']
        local_time.pop(3)
        files = sorted(paths, key=os.path.getctime, reverse=True)
        new_files = []
        for f in files:
            time1 = time.ctime(os.path.getctime(f)).split(' ')
            time1 = [i for i in time1 if i != '']
            time1.pop(3)
            if time1 == local_time and '~$' not in f.name:
                new_files.append(f)
        return new_files if len(new_files) > 0 else [""]

    def the_button_was_clicked_with_choice(self):
        top = tkinter.Tk()
        top.withdraw()
        file_name = fd.askopenfilename(parent=top,
                                       filetypes=(("Word", "*.doc"), ("Word", "*.docx"), ("All files", "*.*")))
        top.destroy()
        if str(file_name).find('.docx') == -1:
            print('Форматирую в .docx...')
            bool_form_docx, new_file_name = self.save_as_docx(os.path.normpath(file_name))
            if bool_form_docx:
                file_name = new_file_name
                print("Готово!")
            else:
                print("Не удалось реформатировать файл :(")
        try:

            frm = Formatter(file_name, settings=self.settings.settings, path_to_save=self.settings.path_to_save)
            frm.Redact()

        except FileNotFoundError:
            print(file_name)
            return
        except Exception as exc:
            print(exc, 'the_button_was_clicked_with_choice')

    def the_button_was_clicked(self):
        file_name = self.get_word()[0]
        if file_name == "":
            print("Сейчас не открыт ни одни документ из базы. Откройте и повторите попытку.")
        if str(file_name).find('.docx') == -1:
            print('Форматирую в .docx...')
            bool_form_docx, new_file_name = self.save_as_docx(os.path.normpath(file_name))
            if bool_form_docx:
                file_name = new_file_name
                print("Готово!")
            else:
                print("Не удалось реформатировать файл :(")
        try:
            frm = Formatter(file_name, settings=self.settings.settings, path_to_save=self.settings.path_to_save)

            frm.Redact()
        except FileNotFoundError:
            return
        except Exception as exc:
            print(exc)

    def open_settings(self):
        self.settings.show()


# Мониторинг новый запущенных программ
app = QApplication(sys.argv)
app.setWindowIcon(QtGui.QIcon('icon_app.png'))
window = MainWindow()
window.setWindowIcon(QtGui.QIcon('icon_app.png'))
app.exec()
