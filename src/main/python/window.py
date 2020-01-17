#!/usr/bin/env python3
import openpyxl
import networkx as nx
from PyQt5 import uic, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QTableWidgetItem

from alphabetspinbox import int2alpha


DEFAULT_COLOR = QtGui.QColor('white')
INTVW_COLOR = QtGui.QColor(240, 198, 116, 100)
TMSLT_COLOR = QtGui.QColor(138, 190, 183, 100)
ERROR_COLOR = QtGui.QColor(204, 102, 102, 100)


class MainWindow(QMainWindow):
    def __init__(self, context):
        super().__init__()
        uic.loadUi(context.get_ui(), self)

        self.interviewees = SelectionRange(self.inputTableWidget,
                                           color=INTVW_COLOR,
                                           update=self.update_interviewees)
        self.timeslots = SelectionRange(self.inputTableWidget,
                                        color=TMSLT_COLOR,
                                        update=self.update_timeslots)

        self.fileLoadButton.clicked.connect(self.open_xlsx)
        self.fileWriteButton.clicked.connect(self.compute_matches)
        self.intvwSpinbox.valueChanged.connect(self.fill_ranges)
        self.tmsltSpinbox.valueChanged.connect(self.fill_ranges)
        self.startRowSpinbox.valueChanged.connect(self.fill_ranges)
        self.endRowSpinbox.valueChanged.connect(self.fill_ranges)

    @pyqtSlot()
    def open_xlsx(self):
        # file dialog for xlsx
        # dialog = QFileDialog(parent=self)
        # dialog.setFileMode(QFileDialog.ExistingFile)
        # dialog.setNameFilter('Spreadsheets (*.xlsx)')
        # if not dialog.exec_():
        #     return
        # xlsx = dialog.selectedFiles()[0]
        xlsx = 'example.xlsx'
        # load xlsx into data
        wb = openpyxl.load_workbook(xlsx)
        ws = wb.active
        self.data = [[str(cell.value or '') for cell in rows]
                     for rows in ws.iter_rows()]
        # populate from data
        rows_c, cols_c = len(self.data), len(self.data[0])
        headers = [int2alpha(n) for n in range(cols_c)]
        self.inputTableWidget.setRowCount(rows_c)
        self.inputTableWidget.setColumnCount(cols_c)
        self.inputTableWidget.setHorizontalHeaderLabels(headers)
        self.statusLabel.setText('載入 %d 列資料。' % rows_c)
        # for row in self.data:
        #     items = [QtGui.QStandardItem(field) for field in row]
        #     self.inputTableWidget.appendRow(items)
        for x, rows in enumerate(self.data):
            for y, cell in enumerate(rows):
                item = QTableWidgetItem(self.data[x][y])
                self.inputTableWidget.setItem(x, y, item)
        self.fill_ranges()

    @pyqtSlot()
    def compute_matches(self):
        col_interviewee = self.intvwSpinbox.value() - 1
        col_timeslot = self.tmsltSpinbox.value() - 1
        print(col_interviewee)
        print(col_timeslot)
        row_range = self.startRowSpinbox.value() - 1, self.endRowSpinbox.value()
        interviewees = [self.inputTableWidget.item(row, col_interviewee).text() for row in range(*row_range)]
        timeslots = [self.inputTableWidget.item(row, col_timeslot).text() for row in range(*row_range)]
        B = nx.Graph()
        for interviewee, timeslot in zip(interviewees, timeslots):
            print(interviewee)
            print(timeslot)
            B.add_node(interviewee, bipartite=0)
            B.add_nodes_from(timeslot.split(', '), bipartite=1)
            B.add_edges_from([(interviewee, t) for t in timeslot.split(', ')])
        matches = nx.bipartite.maximum_matching(B)
        print(matches)
        for interviewee in nx.bipartite.sets(B)[0]:
            print('%s -> %s' % (interviewee, matches[interviewee]))

    @pyqtSlot()
    def fill_ranges(self):
        for range in self.interviewees, self.timeslots:
            for item in range.iterate():
                item.setBackground(DEFAULT_COLOR)
            range.update()
            for item in range.iterate():
                item.setBackground(range.color)
        if self.interviewees.cols == self.timeslots.cols:
            for item in self.interviewees.iterate():
                item.setBackground(ERROR_COLOR)

    def update_interviewees(self):
        col = self.intvwSpinbox.value() - 1
        rows = (self.startRowSpinbox.value() - 1,
                self.endRowSpinbox.value())
        return (col, col + 1), rows

    def update_timeslots(self):
        col = self.tmsltSpinbox.value() - 1
        rows = (self.startRowSpinbox.value() - 1,
                self.endRowSpinbox.value())
        return (col, col + 1), rows


class SelectionRange:
    def __init__(self, table, color, update, cols=None, rows=None):
        self.table = table
        self.color = color
        self.update = update
        self.cols = cols
        self.rows = rows

    def update(self):
        self.cols, self.rows = self.func()

    def iterate(self):
        if not self.cols or not self.rows:
            return []
        for col in range(*self.cols):
            for row in range(*self.rows):
                yield self.table.item(row, col)
