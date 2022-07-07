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
        self.flag = False
        self.setWindowTitle("My App")

        self.cb1 = QCheckBox('Заменить N на №', self)
        self.cb2 = QCheckBox('Заменить дату в шапке', self)
        self.cb3 = QCheckBox('Заменить кавычки', self)
        self.cb4 = QCheckBox('Раскрыть аббревиатуру т.е.', self)
        self.cb5 = QCheckBox('Заменить тире', self)

        if list(self.settings.values())[0]:
            self.cb1.setChecked(True)
        if list(self.settings.values())[1]:
            self.cb2.setChecked(True)
        if list(self.settings.values())[2]:
            self.cb3.setChecked(True)
        if list(self.settings.values())[3]:
            self.cb4.setChecked(True)
        if list(self.settings.values())[4]:
            self.cb5.setChecked(True)

        self.btnEnter = QPushButton("Подтвердить")
        self.btnEnter.setCheckable(True)
        self.btnEnter.clicked.connect(self.enter_setting)

        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.cb1, 0, 0)
        self.grid.addWidget(self.cb2, 0, 1)
        self.grid.addWidget(self.cb3, 1, 0)
        self.grid.addWidget(self.cb4, 1, 1)
        self.grid.addWidget(self.cb5, 2, 0)
        self.grid.addWidget(self.btnEnter, 3, 1)

        self.setLayout(self.grid)

        self.setGeometry(500, 500, 120, 300)

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

        self.show()

    def enter_setting(self):
        self.save_checkbox()
        self.close()

    def save_checkbox(self):
        if self.cb1.isChecked():
            self.settings['ChangeNumber'] = True
        else:
            self.settings['ChangeNumber'] = False

        if self.cb2.isChecked():
            self.settings['ChangeDate'] = True
        else:
            self.settings['ChangeDate'] = False

        if self.cb3.isChecked():
            self.settings['ChangeKavich'] = True
        else:
            self.settings['ChangeKavich'] = False

        if self.cb4.isChecked():
            self.settings['ChangeTN'] = True
        else:
            self.settings['ChangeTN'] = False

        if self.cb5.isChecked():
            self.settings['ChangeTire'] = True
        else:
            self.settings['ChangeTire'] = False

        print(self.settings)

