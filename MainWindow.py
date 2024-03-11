import sys
from PyQt5 import QtWidgets, QtCore


class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)

        self.init_ui()

    def init_ui (self):
        self.setWindowTitle("Excel to SQL Converter")

        file_from_path = QtWidgets.QLineEdit(self)
        file_from_path.setGeometry()
        file_from_path.setReadOnly(True)