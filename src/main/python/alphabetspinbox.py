#!/usr/bin/env python3
from PyQt5.QtWidgets import QSpinBox


def int2alpha(num):
    return chr(ord('A') + min(num, 25))


class AlphabetSpinBox(QSpinBox):
    def __init__(self, parent):
        super().__init__(parent)

    def textFromValue(self, num):
        return int2alpha(num)
