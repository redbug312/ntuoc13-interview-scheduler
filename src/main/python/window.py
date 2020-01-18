#!/usr/bin/env python3
import networkx as nx
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot as slot
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QMainWindow, QFileDialog

from models import SpreadsheetTableModel


DEFAULT_COLOR = QColor('white')
INTVW_COLOR = QColor(240, 198, 116, 100)
TMSLT_COLOR = QColor(138, 190, 183, 100)
ERROR_COLOR = QColor(204, 102, 102, 100)


class MainWindow(QMainWindow):
    def __init__(self, context, parent=None):
        super().__init__(parent)
        uic.loadUi(context.get_ui(), self)

        self.sheets = [SpreadsheetTableModel()]
        self.inputTableView.setModel(self.sheets[0])

    @slot()
    def open_xlsx(self):
        # dialog = QFileDialog(parent=self)
        # dialog.setFileMode(QFileDialog.ExistingFile)
        # dialog.setNameFilter('Spreadsheets (*.xlsx)')
        # if not dialog.exec_():
        #     return
        # xlsx = dialog.selectedFiles()[0]
        xlsx = 'example.xlsx'
        self.sheets[0].populate(xlsx)
        self.update_ranges()
        self.statusLabel.setText('載入 %d 列資料。' % self.sheets[0].rowCount())

    @slot()
    def compute_matching(self):
        B = nx.Graph()
        intvws = self.sheets[0].range('interviewee').iterate()
        tmslts = self.sheets[0].range('timeslot').iterate()
        for interviewee, timeslots in zip(intvws, tmslts):
            B.add_node(interviewee, bipartite=0)
            B.add_nodes_from(timeslots.split(', '), bipartite=1)
            B.add_edges_from([(interviewee, t) for t in timeslots.split(', ')])
        matches = nx.bipartite.maximum_matching(B)
        print(matches)
        for interviewee in nx.bipartite.sets(B)[0]:
            print('%s -> %s' % (interviewee, matches[interviewee]))

    @slot()
    def update_ranges(self):
        rows = self.startRowSpinbox.value(), self.endRowSpinbox.value()
        cols_intvw = (self.intvwSpinbox.value(), ) * 2
        cols_tmslt = (self.tmsltSpinbox.value(), ) * 2
        self.sheets[0].setRange('interviewee', rows, cols_intvw, INTVW_COLOR)
        self.sheets[0].setRange('timeslot', rows, cols_tmslt, TMSLT_COLOR)
