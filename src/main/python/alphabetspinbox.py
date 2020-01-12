#!/usr/bin/env python3
from PyQt5.QtWidgets import QSpinBox


class AlphabetSpinBox(QSpinBox):
    def __init__(self, parent):
        super().__init__(parent)

    def textFromValue(self, num):
        return chr(ord('A') + min(num, 25))
