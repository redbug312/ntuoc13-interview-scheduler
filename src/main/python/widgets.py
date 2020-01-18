#!/usr/bin/env python3
from PyQt5.QtCore import pyqtSignal as signal
from PyQt5.QtWidgets import QSpinBox, QTableView


class AlphabetSpinBox(QSpinBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMaximum(25)  # only A-Y to simplify

    def textFromValue(self, num):
        return chr(64 + num)


class SpreadsheetTableView(QTableView):
    dropFile = signal(str, name='dropFile')

    def __init__(self, parent=None):
        super().__init__(parent)

    def dragEnterEvent(self, event):
        event.accept()

    def dragMoveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if len(urls) == 1:
            url = urls[0].url()
            if url.startswith('file://') and url.endswith('.xlsx'):
                # len('file://') == 7
                self.dropFile.emit(url[7:])
                event.accept()
        event.ignore()
