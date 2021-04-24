from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QMetaType
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication
from main_window import Ui_MainWindow
from test import log_in
import sys


# class start_window(QtWidgets.QMainWindow):
#     def __init__(self):
#         super(start_window, self).__init__()
#         self.start = StartWindow()
#         self.iniUI()
#
#     def iniUI(self):
#         self.start.setupUi(self)
#         log_in_button = self.start.pushButton
#         log_in_button.clicked.connect(self.logging)
#
#     def logging(self):
#         service = log_in()
#         main_app = main_window()
#         main_app.show()


class main_window(QtWidgets.QMainWindow):
    def __init__(self):
        super(main_window, self).__init__()
        self.ui = Ui_MainWindow()
        self.iniUI()

    def iniUI(self):
        self.ui.setupUi(self)


app = QtWidgets.QApplication([])
application = main_window()
application.show()

sys.exit(app.exec())
