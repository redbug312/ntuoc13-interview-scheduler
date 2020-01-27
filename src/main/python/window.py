#!/usr/bin/env python3
import networkx as nx
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot as slot
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QMainWindow, QFileDialog

from models import SpreadsheetTableModel


DEFAULT_COLOR = QColor('white')
INTVW_COLOR = QColor(252, 244, 228)
TMSLT_COLOR = QColor(232, 242, 241)
ERROR_COLOR = QColor(235, 195, 195)


class MainWindow(QMainWindow):
    def __init__(self, context, parent=None):
        super().__init__(parent)
        uic.loadUi(context.ui, self)
        uic.loadUi(context.placeholderUi, self.placeholderFrame)
        self.placeholderFrame.setOverlay(self.inputTableView)
        self.placeholderFrame.setContent(context.excelPixmap, '回應表格未開啟')

        self.sheets = [SpreadsheetTableModel() for _ in range(2)]
        self.inputTableView.setModel(self.sheets[0])
        self.outputTableView.setModel(self.sheets[1])

    @slot()
    @slot(str)
    def openXlsx(self, xlsx=None):
        if xlsx is None:
            # dialog = QFileDialog(parent=self)
            # dialog.setAcceptMode(QFileDialog.AcceptSave)
            # dialog.setFileMode(QFileDialog.ExistingFile)
            # dialog.setNameFilter('Spreadsheets (*.xlsx)')
            # if not dialog.exec_():
            #     return False
            # xlsx = dialog.selectedFiles()[0]
            xlsx = 'example.xlsx'
        self.sheets[0].populate(xlsx)
        # View
        self.placeholderFrame.hide()
        self.updateSheetColumnhead()
        self.updateSheetRanges()
        self.intvwSpinbox.setStyleSheet('background-color: %s' % INTVW_COLOR.name())
        self.tmsltSpinbox.setStyleSheet('background-color: %s' % TMSLT_COLOR.name())
        self.statusbar.showMessage('載入 %d 列資料。' % self.sheets[0].rowCount())

    @slot()
    @slot(str)
    def saveXlsx(self, xlsx=None):
        if xlsx is None:
            # dialog = QFileDialog(parent=self)
            # dialog.setAcceptMode(QFileDialog.AcceptSave)
            # dialog.setFileMode(QFileDialog.AnyFile)
            # dialog.setNameFilter('Spreadsheets (*.xlsx)')
            # if not dialog.exec_():
            #     return False
            # xlsx = dialog.selectedFiles()[0]
            xlsx = 'output.xlsx'
        self.sheets[1].export(xlsx)

    @slot()
    def computeMatching(self):
        intvws = self.sheets[0].range('interviewee')
        tmslts = self.sheets[0].range('timeslot')
        cap = self.capacitySpinbox.value()
        # Calculate maximum-flow
        G = nx.DiGraph()
        for interviewee, timeslots in zip(intvws, tmslts):
            if interviewee is None or timeslots is None:
                continue
            for timeslot in timeslots.split(', '):
                G.add_edge('source', timeslot, capacity=cap)
                G.add_edge(timeslot, interviewee, capacity=1)
            G.add_edge(interviewee, 'sink', capacity=1)
        flow_value, flow_dict = nx.maximum_flow(G, 'source', 'sink')
        # Convert maximum-flow to maximum-matching
        matches = {timeslot: [timeslot] +  # columnhead
                             [interviewee for interviewee, flow
                              in flow_dict[timeslot].items() if flow]
                   for timeslot in G.successors('source')}
        matches['無法排入'] = ['無法排入'] + \
            [interviewee for interviewee in G.predecessors('sink')
             if flow_dict[interviewee]['sink'] == 0]
        self.sheets[1].populate(matches)
        self.sheets[1].setColumnhead(True)
        # Views
        self.tabWidget.addTab(self.outputSheetTab, '排程結果')
        self.tabWidget.setCurrentIndex(1)

    @slot()
    def updateSheetRanges(self):
        rows = self.startRowSpinbox.value(), self.endRowSpinbox.value()
        cols_intvw = (self.intvwSpinbox.value(), ) * 2
        cols_tmslt = (self.tmsltSpinbox.value(), ) * 2
        self.sheets[0].setRange('interviewee', rows, cols_intvw, INTVW_COLOR)
        self.sheets[0].setRange('timeslot', rows, cols_tmslt, TMSLT_COLOR)

    @slot()
    def updateSheetColumnhead(self):
        checked = self.columnheadCheckbox.isChecked()
        self.sheets[0].setColumnhead(checked)
