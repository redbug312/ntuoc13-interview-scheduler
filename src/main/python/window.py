#!/usr/bin/env python3
import openpyxl
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QTableWidgetItem


class MainWindow(QMainWindow):
    def __init__(self, context):
        super().__init__()
        uic.loadUi(context.get_ui(), self)
        self.pushButton.clicked.connect(self.open_xlsx)

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
        self.preview_table.setHorizontalHeaderLabels(self.data[0])
        self.preview_table.setColumnCount(cols_c)
        self.preview_table.setRowCount(rows_c)
        self.filename_label.setText('載入 %d 列資料。' % rows_c)
        for x, rows in enumerate(self.data):
            for y, cell in enumerate(rows):
                item = QTableWidgetItem(self.data[x][y])
                self.preview_table.setItem(x, y, item)
