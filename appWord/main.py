import sys

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton
from tkinter import filedialog as fd
from Formatter import Formatter


# Подкласс QMainWindow для настройки главного окна приложения
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("My App")
        button = QPushButton("Press Me!")
        button.setFixedSize(QSize(100, 50))
        button.setCheckable(True)
        button.clicked.connect(self.the_button_was_clicked)
        self.setFixedSize(QSize(400, 300))
        # Устанавливаем центральный виджет Window.
        self.setCentralWidget(button)

    def the_button_was_clicked(self):
        file_name = fd.askopenfilename(filetypes=(("Word", "*.docx"), ("All files", "*.*")))
        try:
            frm = Formatter(file_name)
            frm.Redact()
        except FileNotFoundError:
            return

app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()
