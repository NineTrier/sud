import sys

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout, QCheckBox
from tkinter import filedialog as fd
from Formatter import Formatter


# Подкласс QMainWindow для настройки главного окна приложения
class Settings(QWidget):

    settings: dict

    def __init__(self, settings):
        super().__init__()

        self.settings = settings

        self.setWindowTitle("My App")

        self.cb1 = QCheckBox('Заменить N на №', self)
        self.cb2 = QCheckBox('Заменить дату в шапке', self)
        self.cb3 = QCheckBox('Заменить кавычки', self)
        self.cb4 = QCheckBox('Раскрыть аббревиатуру т.н.', self)
        self.cb5 = QCheckBox('Заменить тире', self)

        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.cb1, 0, 0)
        self.grid.addWidget(self.cb2, 0, 1)
        self.grid.addWidget(self.cb3, 1, 0)
        self.grid.addWidget(self.cb4, 1, 1)
        self.grid.addWidget(self.cb5, 2, 0)

        self.setLayout(self.grid)

        self.setGeometry(500, 500, 120, 300)

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

        self.show()

