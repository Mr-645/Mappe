import sys
import os
from os.path import dirname, realpath, join
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QTableWidget, QTableWidgetItem, QMainWindow
import pandas as pd
from PyQt5.uic import loadUiType

From_Main, _ = loadUiType(join(dirname(__file__), "app.ui"))


class MainWindow(QMainWindow, From_Main):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        # ====== Functionality code starts here ======
        self.pushButton_4.clicked.connect(self.openfile)
        self.pushButton_3.clicked.connect(self.datahead)

        self.connectButtons()
        self.statusBar.showMessage(">")

    def connectButtons(self):
        self.actionExit.triggered.connect(self.close)

    def openfile(self):
        try:
            path = QFileDialog.getOpenFileName(self, 'Open CSV', os.getenv('HOME'), 'CSV(*.csv)')[0]
            self.all_data = pd.read_csv(path)
        except:
            print("Unexpected error 1:", sys.exc_info()[0])

    def datahead(self):
        num_column = 2
        if num_column == 0:
            num_rows = len(self.all_data.index)
        else:
            num_rows = num_column

        try:
            self.tableWidget.setColumnCount(len(self.all_data.columns))
            self.tableWidget.setRowCount(num_rows)
            self.tableWidget.setHorizontalHeaderLabels(self.all_data.columns)
        except:
            print("Unexpected error 2:", sys.exc_info())

        for i in range(num_rows):
            for j in range(len(self.all_data.columns)):
                self.tableWidget.setItem(
                    i, j, QTableWidgetItem(str(self.all_data.iat[i, j])))

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

# ====== Functionality code ends here ======


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = MainWindow()
    mywindow.show()
    sys.exit(app.exec_())
