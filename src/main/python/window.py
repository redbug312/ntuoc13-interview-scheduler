#!/usr/bin/env python3
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QFileDialog


class MainWindow(QMainWindow):
    def __init__(self, context):
        super().__init__()
        uic.loadUi(context.get_ui(), self)
        self.pushButton.clicked.connect(self.open_file)

    @pyqtSlot()
    def on_click(self):
        print('Hello World')

    @pyqtSlot()
    def open_file(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.selectNameFilter('Spreadsheets (*.xlsx)')
        if dialog.exec_():
            filenames = dialog.selectedFiles()
            print(filenames)
