import os
import sys

from PyQt5 import QtSvg, QtGui, QtCore

import Controller
from PyQt5.QtCore import QTime, QSize
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog, QTableWidgetItem, QHeaderView, \
    QCalendarWidget, QListWidgetItem
import csv
import pandas as pd
import qrcode.image.svg

from Trace import Ui_MainWindow

path = ""
counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks = "", "", "", "", "", "", "", "", "", ""
infoList = []


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Set up User Interface
        QMainWindow.__init__(self, parent=parent)
        self.setupUi(self)
        self.initialiseObject()
        self.materialListItem()

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
        self.backButton4.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))
        self.nextButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(4))
        self.writeButton.setEnabled(False)
        self.writeButton.clicked.connect(lambda: self.writeToCSV())
        self.writeButton.clicked.connect(lambda: self.generateQRCode())

        self.iDateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.sODateEdit.setDateTime(QtCore.QDateTime.currentDateTime())

        self.counterTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.cAddressTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.iNumberTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.aToTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.productTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.quantityTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.signatureTextEdit.textChanged.connect(lambda: self.boolWriteButton())

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
        self.excelTable.setHorizontalHeaderLabels(list(df.columns))
        self.excelTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

        for i in range(rowCount):
            for j in range(columnCount):
                self.excelTable.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

        self.updateCounter()

    def writeToCSV(self):
        global path, counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks, infoList

        counter = self.counterTextEdit.toPlainText()
        cAddress = self.cAddressTextEdit.toPlainText()
        iNumber = self.iNumberTextEdit.toPlainText()
        iDate = self.iDateEdit.date().toPyDate().strftime("%d-%m-%Y")
        aTo = self.aToTextEdit.toPlainText()
        product = self.productTextEdit.toPlainText()
        quantity = self.quantityTextEdit.toPlainText()
        signature = self.signatureTextEdit.toPlainText()
        sODate = self.sODateEdit.date().toPyDate()
        if self.remarksTextEdit.toPlainText() != "":
            remarks = self.remarksTextEdit.toPlainText()
        else:
            remarks = "-"
        infoList = [counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks]

        with open(path[0], 'a') as excelFile:
            writer = csv.writer(excelFile)
            writer.writerow(infoList)

        self.updateCounter()

    def boolWriteButton(self):

        if len(self.counterTextEdit.toPlainText().strip()) != 0 and \
                len(self.cAddressTextEdit.toPlainText().strip()) != 0 and \
                len(self.iNumberTextEdit.toPlainText().strip()) != 0 and \
                len(self.aToTextEdit.toPlainText().strip()) != 0 and \
                len(self.productTextEdit.toPlainText().strip()) != 0 and \
                len(self.quantityTextEdit.toPlainText().strip()) != 0 and \
                len(self.signatureTextEdit.toPlainText().strip()) != 0:
            self.writeButton.setEnabled(True)
        else:
            self.writeButton.setEnabled(False)

    def updateCounter(self):
        global path

        df = pd.read_csv(path[0])
        self.counterTextEdit.setPlainText(str(len(df.index) + 1))

    def generateQRCode(self):
        data = "S/N : " + counter + \
               "\nCompany Address : " + cAddress + \
               "\nInvoice Number : " + iNumber + \
               "\nInvoice Date : " + iDate.strftime("%d-%m-%Y") + \
               "\nAttention To : " + aTo + \
               "\nProduct : " + product + \
               "\nQuantity : " + quantity + \
               "\nSignature : " + signature + \
               "\nSigned off Date : " + sODate.strftime("%d-%m-%Y") + \
               "\nRemarks : " + remarks
        img = qrcode.make(data, image_factory=qrcode.image.svg.SvgPathFillImage)
        saveName = str(counter) + "_" + str(iNumber) + "_" + str(iDate.strftime("%d%m%Y"))
        img.save(saveName + ".svg")
        pixmap = QtGui.QPixmap(saveName + ".svg")
        self.qrCode.setPixmap(pixmap.scaled(150, 150, QtCore.Qt.KeepAspectRatio))
        self.qrCode.show()

    def materialListItem(self):
        item1 = QListWidgetItem("A")
        self.materialList.addItem(item1)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    window = Main()
    window.show()

    sys.exit(app.exec())
