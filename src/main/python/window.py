#!/usr/bin/env python3
import openpyxl
import networkx as nx
from PyQt5 import uic, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QTableWidgetItem

from alphabetspinbox import int2alpha


DEFAULT_COLOR = QtGui.QColor('white')
INTVW_COLOR = QtGui.QColor(240, 198, 116, 100)
TIMSL_COLOR = QtGui.QColor(138, 190, 183, 100)
ERROR_COLOR = QtGui.QColor(204, 102, 102, 100)


class MainWindow(QMainWindow):
    def __init__(self, context):
        super().__init__()
        uic.loadUi(context.get_ui(), self)

        self.interviewees = SelectionRange(self.preview_table, QtGui.QColor(240, 198, 116, 100))
        self.timeslots = SelectionRange(self.preview_table, QtGui.QColor(138, 190, 183, 100))
        self.interviewees.update_callback(self.update_interviewees)
        self.timeslots.update_callback(self.update_timeslots)

        self.pushButton.clicked.connect(self.open_xlsx)
        self.pushButton_2.clicked.connect(self.compute_matches)
        self.interviewee_spinbox.valueChanged.connect(self.fill_ranges)
        self.timeslot_spinbox.valueChanged.connect(self.fill_ranges)
        self.start_row_spinbox.valueChanged.connect(self.fill_ranges)
        self.end_row_spinbox.valueChanged.connect(self.fill_ranges)

    @pyqtSlot()
    def open_xlsx(self):
        # file dialog for xlsx
        dialog = QFileDialog(parent=self)
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter('Spreadsheets (*.xlsx)')
        if not dialog.exec_():
            return
        xlsx = dialog.selectedFiles()[0]
        # xlsx = 'example.xlsx'
        # load xlsx into data
        wb = openpyxl.load_workbook(xlsx)
        ws = wb.active
        self.data = [[str(cell.value or '') for cell in rows]
                     for rows in ws.iter_rows()]
        # populate from data
        rows_c, cols_c = len(self.data), len(self.data[0])
        headers = [int2alpha(n) for n in range(cols_c)]
        self.preview_table.setRowCount(rows_c)
        self.preview_table.setColumnCount(cols_c)
        self.preview_table.setHorizontalHeaderLabels(headers)
        self.filename_label.setText('載入 %d 列資料。' % rows_c)
        for x, rows in enumerate(self.data):
            for y, cell in enumerate(rows):
                item = QTableWidgetItem(self.data[x][y])
                self.preview_table.setItem(x, y, item)
        self.fill_ranges()

    @pyqtSlot()
    def compute_matches(self):
        col_interviewee = self.interviewee_spinbox.value()
        col_timeslot = self.timeslot_spinbox.value()
        row_range = self.start_row_spinbox.value() - 1, self.end_row_spinbox.value()
        interviewees = [self.preview_table.item(row, col_interviewee).text() for row in range(*row_range)]
        timeslots = [self.preview_table.item(row, col_timeslot).text() for row in range(*row_range)]
        B = nx.Graph()
        for interviewee, timeslot in zip(interviewees, timeslots):
            B.add_node(interviewee, bipartite=0)
            B.add_nodes_from(timeslot.split(', '), bipartite=1)
            B.add_edges_from([(interviewee, t) for t in timeslot.split(', ')])
        matches = nx.bipartite.maximum_matching(B)
        print(matches)
        for interviewee in nx.bipartite.sets(B)[0]:
            print('%s -> %s' % (interviewee, matches[interviewee]))

    @pyqtSlot()
    def fill_ranges(self):
        default_color = QtGui.QColor('white')
        for range in self.interviewees, self.timeslots:
            for item in range.iterate():
                item.setBackground(default_color)
            range.update()
            for item in range.iterate():
                item.setBackground(range.color)
        if self.interviewees.cols == self.timeslots.cols:
            for item in self.interviewees.iterate():
                item.setBackground(ERROR_COLOR)

    def update_interviewees(self):
        col = self.interviewee_spinbox.value()
        rows = (self.start_row_spinbox.value() - 1,
                self.end_row_spinbox.value())
        return (col, col + 1), rows

    def update_timeslots(self):
        col = self.timeslot_spinbox.value()
        rows = (self.start_row_spinbox.value() - 1,
                self.end_row_spinbox.value())
        return (col, col + 1), rows


class SelectionRange:
    def __init__(self, table, color, cols=None, rows=None):
        self.table = table
        self.color = color
        self.cols = cols
        self.rows = rows

    def update_callback(self, func):
        self.func = func

    def update(self):
        self.cols, self.rows = self.func()

    def iterate(self):
        if not self.cols or not self.rows:
            return []
        for col in range(*self.cols):
            for row in range(*self.rows):
                yield self.table.item(row, col)
