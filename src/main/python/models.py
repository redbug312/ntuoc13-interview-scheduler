#!/usr/bin/env python3
import openpyxl
import pandas as pd
from datetime import datetime
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QColor

from widgets import AlphabetSpinBox


DEFAULT_COLOR = QColor('white')
ERROR_COLOR = QColor(204, 102, 102, 100)


class SpreadsheetTableModel(QAbstractTableModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ranges = {}
        self.frame = pd.DataFrame()
        self.columnhead = False

    def populate(self, source):
        self.layoutAboutToBeChanged.emit()
        if type(source) is str and source.endswith('.xlsx'):
            ws = openpyxl.load_workbook(source).active
            self.frame = pd.DataFrame(ws.values)
        else:
            self.frame = pd.DataFrame(source)
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

    def columnhead(self):
        return self.columnhead

    def setColumnhead(self, header):
        self.layoutAboutToBeChanged.emit()
        self.columnhead = header
        self.layoutChanged.emit()

    # QAbstractItemModel override functions below

    def sort(self, ncol, order):
        column = self.frame.columns[ncol]
        head = self.frame.head(1)
        self.layoutAboutToBeChanged.emit()
        if self.columnhead:  # take out header before sort
            self.frame.drop(head.index, inplace=True)
        self.frame.sort_values(by=column, ascending=order,
                               kind='mergesort', inplace=True)
        if self.columnhead:  # put back header after sort
            self.frame = pd.concat([head, self.frame])
        self.layoutChanged.emit()

    def flags(self, index):
        flag = super().flags(index)
        if self.columnhead and index.row() == 0:
            flag &= ~Qt.ItemIsEnabled
        return flag

    def rowCount(self, parent=QModelIndex()):
        return self.frame.shape[0]

    def columnCount(self, parent=QModelIndex()):
        return self.frame.shape[1]

    def headerData(self, section, orientation, role):
        if role != Qt.DisplayRole:
            return None
        elif orientation == Qt.Horizontal:
            return AlphabetSpinBox.textFromValue(None, section + 1)
        else:
            return section + 1

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            value = self.frame.iloc[index.row(), index.column()]
            if type(value) in [pd.Timestamp, datetime]:
                return str(value)
            return value
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
