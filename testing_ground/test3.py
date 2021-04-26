# imports
import os
import sys
from os.path import dirname, realpath, join
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QTableWidget, QTableWidgetItem
from PyQt5.uic import loadUiType
import pandas as pd

# load ui file
#baseUIClass, baseUIWidget = uic.loadUiType(join(dirname(__file__), "app.ui"))
From_Main, _ = loadUiType(join(dirname(__file__), "app.ui"))

# use loaded ui file in the logic class
class MainWindow(QWidget, From_Main):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

def main():
    app = QApplication(sys.argv)
    ui = MainWindow()
    ui.show()
    sys.exit(app.exec_())

main()