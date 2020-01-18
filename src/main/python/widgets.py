#!/usr/bin/env python3
from PyQt5.QtWidgets import QSpinBox


class AlphabetSpinBox(QSpinBox):
    def __init__(self, parent=None):
        super().__init__(parent)

    def textFromValue(self, num):
        return chr(64 + min(num, 25))  # assume < 25 columns
