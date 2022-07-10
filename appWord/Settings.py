import json
import sys

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QGridLayout, QCheckBox


# Подкласс QMainWindow для настройки главного окна приложения
class Settings(QWidget):

    settings: dict
    save_setting_file = "settings.set"

    def __init__(self):
        super().__init__()

        if not self.open_save_setting_file():
            self.settings = {'ChangeNumber': True, 'ChangeDate': True, 'ChangeKavich': True, 'ChangeTN': True, 'ChangeTire': True
                             , 'ChangePadeg': True, 'ChangeRF': True, 'ChangeGod': True}
        self.flag = False
        self.setWindowTitle("Настройки")

        self.cb1 = QCheckBox('Заменить N на №', self)
        self.cb2 = QCheckBox('Заменить дату в шапке', self)
        self.cb3 = QCheckBox('Заменить кавычки', self)
        self.cb4 = QCheckBox('Раскрыть аббревиатуру т.е.', self)
        self.cb5 = QCheckBox('Заменить тире', self)
        self.cb6 = QCheckBox('Решение/определение падеж', self)
        self.cb7 = QCheckBox('Раскрыть РФ', self)
        self.cb8 = QCheckBox('Раскрыть год (22 -> 2022)', self)

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
            self.cb7.setChecked(True)
        if list(self.settings.values())[7]:
            self.cb8.setChecked(True)

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
        self.grid.addWidget(self.cb6, 2, 1)
        self.grid.addWidget(self.cb7, 3, 0)
        self.grid.addWidget(self.cb8, 3, 1)
        self.grid.addWidget(self.btnEnter, 4, 1)

        self.setLayout(self.grid)

        self.setGeometry(500, 500, 120, 300)

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

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

        if self.cb7.isChecked():
            self.settings['ChangeGod'] = True
        else:
            self.settings['ChangeGod'] = False

        try:
            with open(self.save_setting_file, 'w') as outfile:
                json.dump(self.settings, outfile)
                outfile.close()
        except:
            print("Не получилось сохранить настройки, попробуйте ещё раз")
        print(self.settings)
