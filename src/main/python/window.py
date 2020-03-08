#!/usr/bin/env python3
import networkx as nx
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot as slot
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox

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
        self.tabWidget.removeTab(1)

        self.sheets = [SpreadsheetTableModel() for _ in range(2)]
        self.inputTableView.setModel(self.sheets[0])
        self.outputTableView.setModel(self.sheets[1])

        self.connects = [sig.connect(slt) for sig, slt in {
            self.fileLoadButton.clicked:       self.openXlsx,
            self.scheduleButton.clicked:       self.schedule,
            self.fileWriteButton.clicked:      self.saveXlsx,
            self.placeholderFrame.dropped:     lambda f: self.openXlsx(f),
            self.inputTableView.dropped:       lambda f: self.openXlsx(f),
            self.columnheadCheckbox.clicked:   lambda: self.updateSpreadsheet(2),
            self.intvwSpinbox.valueChanged:    lambda: self.updateSpreadsheet(4),
            self.tmsltSpinbox.valueChanged:    lambda: self.updateSpreadsheet(4),
            self.startRowSpinbox.valueChanged: lambda: self.updateSpreadsheet(4),
            self.endRowSpinbox.valueChanged:   lambda: self.updateSpreadsheet(4),
            self.outputTableView.selectionModel().selectionChanged: self.selectOutput,
        }.items()]

    @slot()
    @slot(str)
    def openXlsx(self, xlsx=None):
        if xlsx is None:
            dialog = QFileDialog()
            dialog.setAcceptMode(QFileDialog.AcceptOpen)
            dialog.setFileMode(QFileDialog.ExistingFile)
            dialog.setNameFilter('Spreadsheets (*.xlsx)')
            if not dialog.exec_():
                return False
            xlsx = dialog.selectedFiles()[0]
            # xlsx = 'oc12.xlsx'
        self.sheets[0].populate(xlsx)
        # View
        self.placeholderFrame.hide()
        self.updateSpreadsheet()
        self.intvwSpinbox.setStyleSheet('background-color: %s' % INTVW_COLOR.name())
        self.tmsltSpinbox.setStyleSheet('background-color: %s' % TMSLT_COLOR.name())
        self.statusbar.showMessage('載入 %d 列資料。' % self.sheets[0].rowCount())

    @slot()
    def saveXlsx(self, xlsx=None):
        if xlsx is None:
            dialog = QFileDialog()
            dialog.setAcceptMode(QFileDialog.AcceptSave)
            dialog.setFileMode(QFileDialog.AnyFile)
            dialog.setNameFilter('Spreadsheets (*.xlsx)')
            if not dialog.exec_():
                return False
            xlsx = dialog.selectedFiles()[0]
            # xlsx = 'output.xlsx'
        if not xlsx.endswith('.xlsx'):
            xlsx += '.xlsx'
        self.sheets[1].export(xlsx)

    @slot()
    def schedule(self):
        intvws = self.sheets[0].range('interviewee')
        tmslts = self.sheets[0].range('timeslot')
        cap = self.capacitySpinbox.value()
        # Sanitize data from input spreadsheet
        self.inputs = {interviewee: timeslots.split(', ') if timeslots is not None else []
                       for interviewee, timeslots in zip(intvws, tmslts) if interviewee is not None}
        duplicates = find_duplicates(intvws)
        if duplicates:
            QMessageBox.warning(self, '出現重複回應', f'重複的面試人員為：{"、".join(duplicates)}')
        # Calculate maximum-flow
        G = nx.DiGraph()
        for interviewee, timeslots in self.inputs.items():
            for timeslot in timeslots:
                G.add_edge('source', timeslot, capacity=cap)
                G.add_edge(timeslot, interviewee, capacity=1)
            G.add_edge(interviewee, 'sink', capacity=1)
        flow_value, flow_dict = nx.maximum_flow(G, 'source', 'sink')
        # Convert maximum-flow to maximum-matching
        matches = {timeslot: [timeslot] +  # columnhead
                             [interviewee for interviewee, flow
                              in flow_dict[timeslot].items() if flow]
                   for timeslot in G.successors('source')}
        matches['時段為空'] = ['時段為空'] + \
            [interviewee for interviewee in G.predecessors('sink')
             if G.in_degree(interviewee) == 0]
        matches['無法排入'] = ['無法排入'] + \
            [interviewee for interviewee in G.predecessors('sink')
             if G.in_degree(interviewee) != 0 and flow_dict[interviewee]['sink'] == 0]
        self.sheets[1].populate(matches)
        self.sheets[1].setColumnhead(True)
        # Views
        self.statusbar.showMessage(f'排程完成，總共安排 {flow_value} 位面試人員')
        self.tabWidget.addTab(self.outputSheetTab, '排程結果')
        self.tabWidget.setCurrentIndex(1)

    @slot(int)
    def updateSpreadsheet(self, flags=7):
        # Update order determined by the spinboxes read/write operations
        if flags & 1:  # shape of spreadsheet
            cols = self.sheets[0].columnCount()
            self.intvwSpinbox.setMaximum(cols)
            self.tmsltSpinbox.setMaximum(cols)
            rows = self.sheets[0].rowCount()
            self.startRowSpinbox.setMaximum(rows)
            self.endRowSpinbox.setMaximum(rows)
            self.endRowSpinbox.setValue(rows)
        if flags & 2:  # columnhead of spreadsheet
            checked = self.columnheadCheckbox.isChecked()
            self.sheets[0].setColumnhead(checked)
            self.startRowSpinbox.setMinimum(1 + int(checked))
            self.endRowSpinbox.setMinimum(1 + int(checked))
        if flags & 4:  # ranges in spreadsheet
            rows = self.startRowSpinbox.value(), self.endRowSpinbox.value()
            cols_intvw = (self.intvwSpinbox.value(), ) * 2
            cols_tmslt = (self.tmsltSpinbox.value(), ) * 2
            self.sheets[0].setRange('interviewee', rows, cols_intvw, INTVW_COLOR)
            self.sheets[0].setRange('timeslot', rows, cols_tmslt, TMSLT_COLOR)

    @slot()
    def selectOutput(self):
        indices = self.outputTableView.selectedIndexes()
        interviewee = self.outputTableView.model().data(indices[0])
        if interviewee is not None:
            timeslots = '、'.join(self.inputs[interviewee])
            self.statusbar.showMessage(f'{interviewee}：{timeslots}')


def find_duplicates(items):
    seen = set()
    dups = [item for item in items if item in seen or seen.add(item)]
    return dups
