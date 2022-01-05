import sys
import Controller
from PyQt5.QtCore import QTime
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox

from Trace import Ui_MainWindow


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Set up User Interface
        QMainWindow.__init__(self, parent=parent)
        self.setupUi(self)
        self.initialiseObject()


    def initialiseObject(self):
        self.newExcelButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        self.createFileButton.clicked.connect(lambda: Controller.createCSV())



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion")
    window = Main()
    window.show()

    sys.exit(app.exec())
