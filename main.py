import os
import sys
import Controller
from PyQt5.QtCore import QTime
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog, QTableWidgetItem
import csv
import pandas as pd

from Trace import Ui_MainWindow

path = ""


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Set up User Interface
        QMainWindow.__init__(self, parent=parent)
        self.setupUi(self)
        self.initialiseObject()

    def initialiseObject(self):
        self.newExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        self.createFileButton.clicked.connect(lambda: self.createCSV())
        self.createFileButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton3.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        self.backButton3.clicked.connect(lambda: self.excelTable.clear())
        self.importExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        self.openFileDialogButton.clicked.connect(lambda: self.openFileDialog())
        self.openFileButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))
        self.openFileButton.clicked.connect(lambda: self.populateTable())

    def createCSV(self):
        # Get CSV name
        csvFileName = self.newFileNameTextEdit.toPlainText() + ".csv"

        # open the file in the write mode
        with open(csvFileName, 'w') as newFile:
            # create the csv writer
            writer = csv.writer(newFile)
            header = ['S/N', 'Company Address', 'Invoice No.', 'Invoice Date', 'Attn To', 'Model', 'Quantity',
                      'Signature', 'Date']

            # write a row to the csv file
            writer.writerow(header)

        msg = QMessageBox()
        msg.setWindowTitle("Success!")
        msg.setText("File successfully created.")
        msg.exec_()

    def openFileDialog(self):
        global path

        directory = os.path.dirname(__file__)
        path = QFileDialog.getOpenFileName(self, "Import File", directory, 'All Files (*.*)')

        if path[0].endswith(".text") or path[0].endswith(".txt") or path[0].endswith(".csv") or path[0].endswith(
                ".xslx"):
            self.openFileNameTextEdit.setPlainText(path[0])
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Error!")
            msg.setText("Invalid file type. Please select another file.")
            msg.exec_()

    def populateTable(self):
        global path

        df = pd.read_csv(path[0])
        rowCount = len(df.index)
        columnCount = len(df.columns)
        self.excelTable.setColumnCount(columnCount)
        self.excelTable.setRowCount(rowCount)
        self.excelTable.setHorizontalHeaderLabels(df.columns)

        for i in range(rowCount):
            for j in range(columnCount):
                self.excelTable.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    window = Main()
    window.show()

    sys.exit(app.exec())
