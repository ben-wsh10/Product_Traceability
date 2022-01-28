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

csvFileName = ""
materialPath, assembledPath = "", ""
counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks = "", "", "", "", "", "", "", "", "", ""
counter2, pName, tBy, tDate, signature2, sODate2, remarks2, materials = "", "", "", "", "", "", "", ""

partList, assembledList = [], []


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Set up User Interface
        QMainWindow.__init__(self, parent=parent)
        self.setupUi(self)
        self.initialiseObject()

    def initialiseObject(self):

        self.newExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        self.newExcelButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(5))

        self.createFileButton.clicked.connect(lambda: self.createCSV())
        self.createFileButton2.clicked.connect(lambda: self.createCSV())
        self.createFileButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.createFileButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))

        self.backButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton3.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        self.backButton3.clicked.connect(lambda: self.excelTable.clear())
        self.backButton4.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))
        self.backButton5.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(7))
        self.backButton6.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton7.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton8.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(6))

        self.nextButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(4))
        self.nextButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(5))
        self.nextButton3.clicked.connect(lambda: self.getItem())
        self.nextButton4.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(8))
        self.writeButton.setEnabled(False)
        self.writeButton.clicked.connect(lambda: self.writeToCSV())
        self.writeButton2.clicked.connect(lambda: self.writeToCSV())
        self.writeButton.clicked.connect(lambda: self.generateQRCode())
        self.writeButton2.clicked.connect(lambda: self.generateQRCode())

        self.importExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        self.importExcelButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(6))

        self.openFileDialogButton.clicked.connect(lambda: self.openMaterialFileDialog())
        self.openFileDialogButton2.clicked.connect(lambda: self.openMaterialFileDialog())
        self.openFileDialogButton3.clicked.connect(lambda: self.openAssembledFileDialog())

        self.openFileButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))
        self.openFileButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(7))
        self.openFileButton.clicked.connect(lambda: self.populateTable())
        self.openFileButton2.clicked.connect(lambda: self.populateTable())

        self.iDateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.tDateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.sODateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.sODateEdit2.setDateTime(QtCore.QDateTime.currentDateTime())

        self.counterTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.cAddressTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.iNumberTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.aToTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.productTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.quantityTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.signatureTextEdit.textChanged.connect(lambda: self.boolWriteButton())

    def createCSV(self):
        # Get CSV name
        global csvFileName

        if self.stackedWidget.currentIndex() == 1:
            csvFileName = self.newFileNameTextEdit.toPlainText() + ".csv"
        elif self.stackedWidget.currentIndex() == 5:
            csvFileName = self.newFileNameTextEdit2.toPlainText() + ".csv"

        # open the file in the write mode
        with open(csvFileName, 'w') as newFile:
            # create the csv writer
            writer = csv.writer(newFile)
            if self.stackedWidget.currentIndex() == 1:
                header = ['S/N', 'Company Address', 'Invoice No.', 'Invoice Date', 'Attn To', 'Model', 'Quantity',
                          'Signature', 'Date']
            elif self.stackedWidget.currentIndex() == 5:
                header = ['S/N', 'Product Name', 'Tested By', 'Testing Date', 'Signature', 'Sign off Date',
                          'Remarks', 'Materials']

            # write a row to the csv file
            writer.writerow(header)

        msg = QMessageBox()
        msg.setWindowTitle("Success!")
        msg.setText("File successfully created.")
        msg.exec_()

    def openMaterialFileDialog(self):
        global materialPath

        directory = os.path.dirname(__file__)
        materialPath = QFileDialog.getOpenFileName(self, "Import File", directory, 'All Files (*.*)')

        if materialPath[0].endswith(".text") or materialPath[0].endswith(".txt") or materialPath[0].endswith(".csv") or \
                materialPath[0].endswith(".xslx"):
            if self.stackedWidget.currentIndex() == 2:
                self.openFileNameTextEdit.setPlainText(materialPath[0])
            elif self.stackedWidget.currentIndex() == 6:
                self.openFileNameTextEdit2.setPlainText(materialPath[0])
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Error!")
            msg.setText("Invalid file type. Please select another file.")
            msg.exec_()

    def openAssembledFileDialog(self):
        global assembledPath

        directory = os.path.dirname(__file__)
        assembledPath = QFileDialog.getOpenFileName(self, "Import File", directory, 'All Files (*.*)')

        if assembledPath[0].endswith(".text") or assembledPath[0].endswith(".txt") or assembledPath[0].endswith(
                ".csv") or assembledPath[0].endswith(".xslx"):
            if self.stackedWidget.currentIndex() == 6:
                self.openFileNameTextEdit3.setPlainText(assembledPath[0])
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Error!")
            msg.setText("Invalid file type. Please select another file.")
            msg.exec_()

    def populateTable(self):
        global materialPath

        if self.stackedWidget.currentIndex() == 3:
            df = pd.read_csv(materialPath[0])
            rowCount = len(df.index)
            columnCount = len(df.columns)
            self.excelTable.setColumnCount(columnCount)
            self.excelTable.setRowCount(rowCount)
            self.excelTable.setHorizontalHeaderLabels(list(df.columns))
            self.excelTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

            for i in range(rowCount):
                for j in range(columnCount):
                    self.excelTable.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

        elif self.stackedWidget.currentIndex() == 7:
            df = pd.read_csv(assembledPath[0])
            rowCount = len(df.index)
            columnCount = len(df.columns)
            self.excelTable2.setColumnCount(columnCount)
            self.excelTable2.setRowCount(rowCount)
            self.excelTable2.setHorizontalHeaderLabels(list(df.columns))
            self.excelTable2.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

            for i in range(rowCount):
                for j in range(columnCount):
                    self.excelTable2.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

        self.updateCounter()

        self.materialListItem()

    def writeToCSV(self):
        global materialPath, counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks, \
            assembledPath, counter2, pName, tBy, tDate, signature2, sODate2, remarks2, materials, partList, assembledList

        if self.stackedWidget.currentIndex() == 4:
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
            partList = [counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks]

            with open(materialPath[0], 'a') as excelFile:
                writer = csv.writer(excelFile)
                writer.writerow(partList)

        elif self.stackedWidget.currentIndex() == 8:
            counter2 = self.counterTextEdit2.toPlainText()
            pName = self.pNameTextEdit.toPlainText()
            tBy = self.tByTextEdit.toPlainText()
            tDate = self.tDateEdit.date().toPyDate().strftime("%d-%m-%Y")
            signature2 = self.signatureTextEdit2.toPlainText()
            sODate2 = self.sODateEdit2.date().toPyDate().strftime("%d-%m-%Y")
            if self.remarksTextEdit2.toPlainText():
                remarks2 = self.remarksTextEdit2.toPlainText()
            else:
                remarks2 = "-"

            materials = self.getItem()

            assembledList = [counter2, pName, tBy, tDate, signature2, sODate2, materials]

            with open(assembledPath[0], 'a+', newline="") as excelFile:
                writer = csv.writer(excelFile)
                writer.writerow(assembledList)

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
        global materialPath, assembledPath

        if self.stackedWidget.currentIndex() == 3:
            df = pd.read_csv(materialPath[0])
            self.counterTextEdit.setPlainText(str(len(df.index) + 1))
        elif self.stackedWidget.currentIndex() == 7:
            df = pd.read_csv(assembledPath[0])
            self.counterTextEdit2.setPlainText(str(len(df.index) + 1))

    def generateQRCode(self):

        if self.stackedWidget.currentIndex() == 4:
            data = "S/N : " + counter + \
                   "\nCompany Address : " + cAddress + \
                   "\nInvoice Number : " + iNumber + \
                   "\nInvoice Date : " + str(iDate) + \
                   "\nAttention To : " + aTo + \
                   "\nProduct : " + product + \
                   "\nQuantity : " + quantity + \
                   "\nSignature : " + signature + \
                   "\nSigned off Date : " + str(sODate) + \
                   "\nRemarks : " + remarks

            img = qrcode.make(data, image_factory=qrcode.image.svg.SvgPathFillImage)
            saveName = str(counter) + "_" + str(iNumber) + "_" + str(iDate.replace("-", ""))
            img.save(saveName + ".svg")
            pixmap = QtGui.QPixmap(saveName + ".svg")
            self.qrCode.setPixmap(pixmap.scaled(150, 150, QtCore.Qt.KeepAspectRatio))
            self.qrCode.show()
        elif self.stackedWidget.currentIndex() == 8:
            # data = "S/N : " + counter2 +"\nProduct Name : " + pName +"\nTested By : " + tBy + "\nTesting Date : " + tDate + "\nSigned off Date : " + sODate2 + "\nRemarks : " + remarks2
            data = "S/N : " + counter2 + \
                   "\nProduct Name : " + pName + \
                   "\nTested By : " + tBy + \
                   "\nTesting Date : " + tDate + \
                   "\nSignature : " + signature2 + \
                   "\nSigned off Date : " + sODate2 + \
                   "\nRemarks : " + remarks2 + \
                   "\nProduct Parts : "
            data = data + self.getItem()
            print(data)
            img = qrcode.make(data, image_factory=qrcode.image.svg.SvgPathFillImage)
            saveName = str(counter) + "_" + str(pName) + "_" + str(tDate.replace("-", ""))
            img.save(saveName + ".svg")
            pixmap = QtGui.QPixmap(saveName + ".svg")
            self.qrCode2.setPixmap(pixmap.scaled(150, 150, QtCore.Qt.KeepAspectRatio))
            self.qrCode2.show()

    def materialListItem(self):
        global materialPath

        df = pd.read_csv(materialPath[0])
        excelList = df.values.tolist()

        for _ in range(len(excelList)):
            tmp = excelList[_]
            convertedList = [str(element) for element in tmp]
            joinedString = ", ".join(convertedList)
            self.materialList.addItem(joinedString)

    def getItem(self):

        extraList = []
        for _ in range(self.includeList.count()):
            tmpCurrentMaterialList = self.includeList.item(_).text().split(",")
            tmpSortedList = ["[" +
                             tmpCurrentMaterialList.pop(5),  # Product Part name
                             tmpCurrentMaterialList.pop(1),  # Company Address
                             tmpCurrentMaterialList.pop(1),  # Invoice Number
                             tmpCurrentMaterialList.pop(1)
                             + " ]"]  # Invoice Date
            tmpString = ",".join(tmpSortedList)
            extraList.append(tmpString)
        extraString = ", ".join(extraList)
        return extraString


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    window = Main()
    window.show()

    sys.exit(app.exec())
