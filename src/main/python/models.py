#!/usr/bin/env python3
import openpyxl
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QColor

from widgets import AlphabetSpinBox


DEFAULT_COLOR = QColor('white')
ERROR_COLOR = QColor(204, 102, 102, 100)


class SpreadsheetTableModel(QAbstractTableModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ranges = {}
        self.sheet = [[]]
        self.row_count = 0
        self.column_count = 0

    def populate(self, source):
        if type(source) is list:
            self.sheet = source
        elif type(source) is str and source.endswith('.xlsx'):
            ws = openpyxl.load_workbook(source).active
            self.sheet = [[str(cell.value or '') for cell in rows]
                          for rows in ws.iter_rows()]
        else:
            return
        self.row_count = len(self.sheet)
        self.column_count = len(self.sheet[0])  # must be header
        self.layoutChanged.emit()

    def range(self, name):
        return self.ranges[name]  # may raise KeyError

    def setRange(self, name, rows, cols, color):
        if name in self.ranges:
            range = self.ranges[name]
            self.dataChanged.emit(*range.corners())
        range = SpreadsheetRange(self, rows, cols, color)
        self.dataChanged.emit(*range.corners())
        self.ranges[name] = range

    def rowCount(self, parent=QModelIndex()):
        return self.row_count

    def columnCount(self, parent=QModelIndex()):
        return self.column_count

    def headerData(self, section, orientation, role):
        if role != Qt.DisplayRole:
            return None
        elif orientation == Qt.Horizontal:
            return AlphabetSpinBox.textFromValue(None, section + 1)
        else:
            return section + 1

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            row, col = index.row(), index.column()
            return self.sheet[row][col]
        elif role == Qt.BackgroundRole:
            includes = [range for range in self.ranges.values()
                        if range.include(index)]
            if len(includes) == 0:
                return DEFAULT_COLOR
            elif len(includes) == 1:
                return includes[0].color
            else:
                return ERROR_COLOR
        elif role == Qt.TextAlignmentRole:
            return Qt.AlignLeft | Qt.AlignVCenter
        return None


class SpreadsheetRange:
    def __init__(self, sheet, rows=(0, 0), cols=(0, 0), color=None):
        self.sheet = sheet
        self.rows = rows[0] - 1, rows[1]
        self.cols = cols[0] - 1, cols[1]
        self.color = color

    def __iter__(self):
        for index in [self.sheet.index(row, col)
                      for row in range(*self.rows)
                      for col in range(*self.cols)]:
            yield self.sheet.data(index)

    def include(self, index):
        return self.rows[0] <= index.row() < self.rows[1] and \
               self.cols[0] <= index.column() < self.cols[1]

    def corners(self):
        return [self.sheet.index(*pos) for pos in zip(self.rows, self.cols)]
