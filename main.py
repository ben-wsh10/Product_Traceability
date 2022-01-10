import sys
import Controller
from PyQt5.QtCore import QTime
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import csv
from Trace import Ui_MainWindow


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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    window = Main()
    window.show()

    sys.exit(app.exec())
