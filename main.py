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
browseMaterial, browseAssembled = False, False
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

        #
        self.newExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        self.newExcelButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(5))
        #
        self.createFileButton.clicked.connect(lambda: self.createCSV())
        self.createFileButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.createFileButton2.clicked.connect(lambda: self.createCSV())
        self.createFileButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        #
        self.backButton.clicked.connect(lambda: self.resetFunction())
        self.backButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton2.clicked.connect(lambda: self.resetFunction())
        self.backButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton3.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        self.backButton3.clicked.connect(lambda: self.excelTable.clear())
        self.backButton4.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))
        self.backButton5.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(7))
        self.backButton6.clicked.connect(lambda: self.resetFunction())
        self.backButton6.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton7.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.backButton8.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(6))
        #
        self.nextButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(4))
        self.nextButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(5))
        self.nextButton3.clicked.connect(lambda: self.getItem())
        self.nextButton4.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(8))
        #
        self.writeButton.setEnabled(False)
        self.writeButton.clicked.connect(lambda: self.writeToCSV())
        self.writeButton.clicked.connect(lambda: self.generateQRCode())
        self.writeButton2.setEnabled(False)
        self.writeButton2.clicked.connect(lambda: self.writeToCSV())
        self.writeButton2.clicked.connect(lambda: self.generateQRCode())
        #
        self.importExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        self.importExcelButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(6))
        #
        self.openFileDialogButton.clicked.connect(lambda: self.openMaterialFileDialog())
        self.openFileDialogButton2.clicked.connect(lambda: self.openMaterialFileDialog())
        self.openFileDialogButton3.clicked.connect(lambda: self.openAssembledFileDialog())
        #
        self.openFileButton.setEnabled(False)
        self.openFileButton2.setEnabled(False)
        self.openFileButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))
        self.openFileButton2.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(7))
        self.openFileButton.clicked.connect(lambda: self.populateTable())
        self.openFileButton2.clicked.connect(lambda: self.populateTable())
        #
        self.iDateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.tDateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.sODateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.sODateEdit2.setDateTime(QtCore.QDateTime.currentDateTime())
        #
        self.counterTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.cAddressTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.iNumberTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.aToTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.productTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.quantityTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.signatureTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        #
        self.counterTextEdit2.textChanged.connect(lambda: self.boolWriteButton())
        self.pNameTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.tByTextEdit.textChanged.connect(lambda: self.boolWriteButton())
        self.signatureTextEdit2.textChanged.connect(lambda: self.boolWriteButton())

    # Reset field, functionality, etc.. where applicable
    def resetFunction(self):
        global browseMaterial, browseAssembled

        if self.stackedWidget.currentWidget().objectName() == "CreateMaterialExcelPage":
            self.newFileNameTextEdit.clear()
        elif self.stackedWidget.currentWidget().objectName() == "CreateAssembledExcelPage":
            self.newFileNameTextEdit2.clear()
        elif self.stackedWidget.currentWidget().objectName() == "ImportMaterialPage":
            self.openFileNameTextEdit.clear()
            self.openFileButton.setEnabled(False)
        elif self.stackedWidget.currentWidget().objectName() == "ImportProductPage":
            self.openFileNameTextEdit2.clear()
            self.openFileNameTextEdit3.clear()
            self.openFileButton2.setEnabled(False)
            browseMaterial, browseAssembled = False, False

    # Create new csv file
    def createCSV(self):
        global csvFileName

        if self.stackedWidget.currentWidget().objectName() == "CreateMaterialExcelPage":
            csvFileName = self.newFileNameTextEdit.toPlainText() + ".csv"
        elif self.stackedWidget.currentWidget().objectName() == "CreateAssembledExcelPage":
            csvFileName = self.newFileNameTextEdit2.toPlainText() + ".csv"

        with open(csvFileName, 'w') as newFile:
            writer = csv.writer(newFile)
            if self.stackedWidget.currentWidget().objectName() == "CreateMaterialExcelPage":
                header = ['S/N', 'Company Address', 'Invoice No.', 'Invoice Date', 'Attn To', 'Model', 'Quantity',
                          'Signature', 'Date']
            elif self.stackedWidget.currentWidget().objectName() == "CreateAssembledExcelPage":
                header = ['S/N', 'Product Name', 'Tested By', 'Testing Date', 'Signature', 'Sign off Date',
                          'Remarks', 'Materials']
            writer.writerow(header)

        msg = QMessageBox()
        msg.setWindowTitle("Success!")
        msg.setText("File successfully created.")
        msg.exec_()

    # Open File Dialog for Material excel sheet
    def openMaterialFileDialog(self):
        global materialPath, browseMaterial

        directory = os.path.dirname(__file__)
        materialPath = QFileDialog.getOpenFileName(self, "Import File", directory, 'All Files (*.*)')

        if materialPath[0].endswith(".text") or materialPath[0].endswith(".txt") or materialPath[0].endswith(".csv") or \
                materialPath[0].endswith(".xlsx"):
            if self.stackedWidget.currentWidget().objectName() == "ImportMaterialPage":
                self.openFileNameTextEdit.setPlainText(materialPath[0])
                self.openFileButton.setEnabled(True)
            elif self.stackedWidget.currentWidget().objectName() == "ImportProductPage":
                self.openFileNameTextEdit2.setPlainText(materialPath[0])
                browseMaterial = True
            if browseMaterial and browseAssembled is True:
                self.openFileButton2.setEnabled(True)
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Error!")
            msg.setText("Invalid file type. Please select another file.")
            msg.exec_()

    # Open File Dialog for Assembled Product excel sheet
    def openAssembledFileDialog(self):
        global assembledPath, browseAssembled

        directory = os.path.dirname(__file__)
        assembledPath = QFileDialog.getOpenFileName(self, "Import File", directory, 'All Files (*.*)')

        if assembledPath[0].endswith(".text") or assembledPath[0].endswith(".txt") or assembledPath[0].endswith(
                ".csv") or assembledPath[0].endswith(".xlsx"):
            if self.stackedWidget.currentWidget().objectName() == "ImportProductPage":
                self.openFileNameTextEdit3.setPlainText(assembledPath[0])
                browseAssembled = True
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Error!")
            msg.setText("Invalid file type. Please select another file.")
            msg.exec_()
        if browseMaterial and browseAssembled is True:
            self.openFileButton2.setEnabled(True)

    # Populate excel sheet table into Widget
    def populateTable(self):
        global materialPath

        if self.stackedWidget.currentWidget().objectName() == "readMaterialExcelPage":
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

        elif self.stackedWidget.currentWidget().objectName() == "readAssembledExcelPage":
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

    # Write input into existing csv files
    def writeToCSV(self):
        global materialPath, counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks, \
            assembledPath, counter2, pName, tBy, tDate, signature2, sODate2, remarks2, materials, partList, assembledList

        if self.stackedWidget.currentWidget().objectName() == "materialInputPage":
            counter = self.counterTextEdit.toPlainText()
            cAddress = self.cAddressTextEdit.toPlainText()
            iNumber = self.iNumberTextEdit.toPlainText()
            iDate = self.iDateEdit.date().toPyDate().strftime("%d-%m-%Y")
            aTo = self.aToTextEdit.toPlainText()
            product = self.productTextEdit.toPlainText()
            quantity = self.quantityTextEdit.toPlainText()
            signature = self.signatureTextEdit.toPlainText()
            sODate = self.sODateEdit.date().toPyDate()
            if self.remarksTextEdit.toPlainText():
                remarks = self.remarksTextEdit.toPlainText()
            else:
                remarks = "-"

            partList = [counter, cAddress, iNumber, iDate, aTo, product, quantity, signature, sODate, remarks]

            with open(materialPath[0], 'w', newline="") as excelFile:
                if partList:
                    writer = csv.writer(excelFile)
                    writer.writerow(partList)

        elif self.stackedWidget.currentWidget().objectName() == "assembledInputPage":
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
            assembledList = [counter2, pName, tBy, tDate, signature2, sODate2, remarks2, materials]

            with open(assembledPath[0], 'a', newline="") as excelFile:
                if assembledList:
                    writer = csv.writer(excelFile)
                    writer.writerow(assembledList)

        self.updateCounter()

    # Condition trigger for writing to csv and qr generator
    def boolWriteButton(self):

        if self.stackedWidget.currentWidget().objectName() == "materialInputPage":
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
        elif self.stackedWidget.currentWidget().objectName() == "assembledInputPage":
            if len(self.counterTextEdit2.toPlainText().strip()) != 0 and \
                    len(self.pNameTextEdit.toPlainText().strip()) != 0 and \
                    len(self.tByTextEdit.toPlainText().strip()) != 0 and \
                    len(self.signatureTextEdit2.toPlainText().strip()) != 0:
                self.writeButton2.setEnabled(True)
            else:
                self.writeButton2.setEnabled(False)

    # Update S/N counter
    def updateCounter(self):
        global materialPath, assembledPath

        if self.stackedWidget.currentWidget().objectName() == "readMaterialExcelPage" or \
            self.stackedWidget.currentWidget().objectName() == "materialInputPage":
            df = pd.read_csv(materialPath[0])
            self.counterTextEdit.setPlainText(str(len(df.index) + 1))
        elif self.stackedWidget.currentWidget().objectName() == "readAssembledExcelPage" or \
            self.stackedWidget.currentWidget().objectName() == "assembledInputPage":
            df = pd.read_csv(assembledPath[0])
            self.counterTextEdit2.setPlainText(str(len(df.index) + 1))

    # QR Code generator
    def generateQRCode(self):
        if self.stackedWidget.currentWidget().objectName() == "materialInputPage":
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
        elif self.stackedWidget.currentWidget().objectName() == "assembledInputPage":
            data = "S/N : " + counter2 + \
                   "\nProduct Name : " + pName + \
                   "\nTested By : " + tBy + \
                   "\nTesting Date : " + tDate + \
                   "\nSignature : " + signature2 + \
                   "\nSigned off Date : " + sODate2 + \
                   "\nRemarks : " + remarks2 + \
                   "\nProduct Parts : "
            # Add Material List items into data
            data = data + self.getItem()
            img = qrcode.make(data, image_factory=qrcode.image.svg.SvgPathFillImage)
            saveName = str(counter) + "_" + str(pName) + "_" + str(tDate.replace("-", ""))
            img.save(saveName + ".svg")
            pixmap = QtGui.QPixmap(saveName + ".svg")
            self.qrCode2.setPixmap(pixmap.scaled(150, 150, QtCore.Qt.KeepAspectRatio))
            self.qrCode2.show()

    # Display Material List items into Widget
    def materialListItem(self):
        global materialPath

        df = pd.read_csv(materialPath[0])
        excelList = df.values.tolist()

        for _ in range(len(excelList)):
            tmp = excelList[_]
            convertedList = [str(element) for element in tmp]
            joinedString = ", ".join(convertedList)
            self.materialList.addItem(joinedString)

    # Return a string of material list items that are included in QR Code
    def getItem(self):
        extraList = []
        for _ in range(self.includeList.count()):
            tmpCurrentMaterialList = self.includeList.item(_).text().split(",")
            tmpSortedList = ["[" +
                             tmpCurrentMaterialList.pop(5),  # Product Part name
                             tmpCurrentMaterialList.pop(1),  # Company Address
                             tmpCurrentMaterialList.pop(1),  # Invoice Number
                             tmpCurrentMaterialList.pop(1)   # Invoice Date
                             + " ]"]
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
