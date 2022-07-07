import os
import sys
import time
from pathlib import Path

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout

from Formatter import Formatter
from Settings import Settings


# Подкласс QMainWindow для настройки главного окна приложения
class MainWindow(QWidget):
    settings = {'ChangeNumber': True, 'ChangeDate': True, 'ChangeKavich': True, 'ChangeTN': True, 'ChangeTire': True}

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

    def get_word(self):
        paths = sorted(Path(f"{os.getenv('LOCALAPPDATA')}\\Temp").glob('*.docx'))
        print(paths)
        local_time = time.ctime(time.time()).split(' ')
        local_time = [i for i in local_time if i != '']
        local_time.pop(3)
        files = sorted(paths, key=os.path.getctime, reverse=True)
        new_files = []
        for f in files:
            time1 = time.ctime(os.path.getctime(f)).split(' ')
            time1 = [i for i in time1 if i != '']
            time1.pop(3)
            if time1 == local_time:
                new_files.append(f)
        return new_files

    def the_button_was_clicked(self):
        #top = tkinter.Tk()
        #top.withdraw()
        #file_name = fd.askopenfilename(parent=top, filetypes=(("Word", "*.docx"), ("All files", "*.*")))
        #top.destroy()
        file_name = self.get_word()[0]
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

