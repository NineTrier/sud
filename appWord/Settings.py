import json
import os

import tkinter

from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout, QCheckBox, QLineEdit, QLabel, \
    QHBoxLayout, QVBoxLayout
from tkinter import filedialog as fd


# Подкласс QMainWindow для настройки главного окна приложения
class Settings(QWidget):
    settings: dict
    save_setting_file = f"{os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')}\\settings.set"

    def __init__(self):
        super().__init__()

        self.default = f"{os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')}\\Ассистент"

        if not self.open_save_setting_file():
            self.settings = {'ChangeNumber': True, 'ChangeDate': True, 'ChangeKavich': True, 'ChangeTN': True,
                             'ChangeTire': True, 'ChangePadeg': True, 'ChangeRF': False, 'ChangeGod': True, 'ChangeTime': True,
                             'ChangeHighlight': False, 'PathToSave': self.default}
        self.flag = False
        self.path_to_save = self.settings['PathToSave']

        self.setWindowTitle("Настройки")

        self.cb1 = QCheckBox('Заменить N на №', self)
        self.cb2 = QCheckBox('Раскрыть дату', self)
        self.cb3 = QCheckBox('Заменить кавычки', self)
        self.cb4 = QCheckBox('Раскрыть аббревиатуру', self)
        self.cb5 = QCheckBox('Заменить тире', self)
        self.cb6 = QCheckBox('Решение/определение падеж', self)
        self.cb7 = QCheckBox('Раскрыть РФ', self)
        self.cb8 = QCheckBox('Раскрыть год (22 -> 2022)', self)
        self.cb9 = QCheckBox('Раскрыть время (часов, минут)', self)
        self.cb10 = QCheckBox('Убрать жёлтый маркер', self)

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
        if list(self.settings.values())[5]:
            self.cb6.setChecked(True)
        if list(self.settings.values())[6]:
            self.cb7.setChecked(False)
        if list(self.settings.values())[7]:
            self.cb8.setChecked(True)
        if list(self.settings.values())[8]:
            self.cb9.setChecked(True)
        if list(self.settings.values())[9]:
            self.cb10.setChecked(False)

        self.label1 = QLabel("Путь:")

        self.input1 = QLineEdit()
        self.input1.setText(str(self.path_to_save))
        self.input1.setReadOnly(True)

        self.btnPath = QPushButton("Выбрать")
        self.btnPath.setCheckable(True)
        self.btnPath.clicked.connect(self.set_path_to_save)

        self.btnEnter = QPushButton("Подтвердить")
        self.btnEnter.setCheckable(True)
        self.btnEnter.clicked.connect(self.enter_setting)

        self.btnDefault = QPushButton("По умолчанию")
        self.btnDefault.setCheckable(True)
        self.btnDefault.clicked.connect(self.set_default)

        self.hbox = QHBoxLayout()
        self.hbox.setSpacing(5)
        self.hbox.addWidget(self.label1)
        self.hbox.addWidget(self.input1)
        self.hbox.addWidget(self.btnPath)

        self.hboxBottom = QHBoxLayout()
        self.hboxBottom.setSpacing(5)
        self.hboxBottom.addWidget(self.btnDefault)
        self.hboxBottom.addWidget(self.btnEnter)

        self.grid = QGridLayout()
        self.grid.setSpacing(10)
        self.grid.addWidget(self.cb1, 0, 0)
        self.grid.addWidget(self.cb2, 0, 1)
        self.grid.addWidget(self.cb3, 1, 0)
        self.grid.addWidget(self.cb4, 1, 1)
        self.grid.addWidget(self.cb5, 2, 0)
        self.grid.addWidget(self.cb6, 2, 1)
        self.grid.addWidget(self.cb7, 3, 0)
        self.grid.addWidget(self.cb8, 3, 1)
        self.grid.addWidget(self.cb9, 4, 0)
        self.grid.addWidget(self.cb10, 4, 1)

        self.vbox = QVBoxLayout()
        self.vbox.setSpacing(5)
        self.vbox.addLayout(self.grid)
        self.vbox.addLayout(self.hbox)
        self.vbox.addLayout(self.hboxBottom)

        self.setLayout(self.vbox)

        self.setGeometry(500, 500, 120, 300)

    def set_path_to_save(self):
        top = tkinter.Tk()
        top.withdraw()
        file_name = fd.askdirectory()
        if file_name != "":
            file_name += '/Ассистент'
            self.settings['PathToSave'] = file_name
            self.input1.setText(str(file_name))
            self.path_to_save = str(file_name)
        top.destroy()

    def open_save_setting_file(self):
        try:
            with open(self.save_setting_file, "r") as write_file:
                setting_file = json.load(write_file)
                self.settings = setting_file
                write_file.close()
                return True
        except:
            return False

    def enter_setting(self):
        self.save_checkbox()
        self.close()

    def set_default(self):
        self.settings = {'ChangeNumber': True, 'ChangeDate': True, 'ChangeKavich': True, 'ChangeTN': True,
                         'ChangeTire': True, 'ChangePadeg': True, 'ChangeRF': False, 'ChangeGod': True, 'ChangeTime': True,
                         'ChangeHighlight': False, 'PathToSave': self.default}
        self.save_checkbox()
        self.path_to_save = self.default
        self.input1.setText(self.default)

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

        if self.cb6.isChecked():
            self.settings['ChangePadeg'] = True
        else:
            self.settings['ChangePadeg'] = False

        if self.cb7.isChecked():
            self.settings['ChangeRF'] = True
        else:
            self.settings['ChangeRF'] = False

        if self.cb8.isChecked():
            self.settings['ChangeGod'] = True
        else:
            self.settings['ChangeGod'] = False

        if self.cb9.isChecked():
            self.settings['ChangeTime'] = True
        else:
            self.settings['ChangeTime'] = False

        if self.cb10.isChecked():
            self.settings['ChangeHighlight'] = True
        else:
            self.settings['ChangeHighlight'] = False

        try:
            with open(self.save_setting_file, 'w') as outfile:
                json.dump(self.settings, outfile)
                outfile.close()
        except:
            print("Не получилось сохранить настройки, попробуйте ещё раз")
        print(self.settings)
