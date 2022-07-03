import sys

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout
from tkinter import filedialog as fd
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
        self.btnOpenFile.clicked.connect(self.open_settings)

        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.btnOpenFile, 0, 0)
        self.grid.addWidget(self.btnFormat, 1, 0)
        self.grid.addWidget(self.btnSettings, 2, 0)

        self.setLayout(self.grid)

        self.setGeometry(app.screens()[0].size().width()-120, (app.screens()[0].size().height()//2)-150, 120, 300)

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

        self.show()


    def the_button_was_clicked(self):
        file_name = fd.askopenfilename(filetypes=(("Word", "*.docx"), ("All files", "*.*")))
        try:
            frm = Formatter(file_name, settings=self.settings)
            frm.Redact()
        except FileNotFoundError:
            return

    def open_settings(self):
        settings_window = Settings(self.settings)


app = QApplication(sys.argv)
window = MainWindow()

app.exec()
