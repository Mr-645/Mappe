import sys
import os
from os.path import dirname, realpath, join
from PyQt5.QtWidgets import QApplication, QDialog, QWidget, QFileDialog, QTableWidget, QTableWidgetItem, QMainWindow, QMessageBox
import json
from PyQt5.uic import loadUi

from app_latest import Ui_MainWindow


class Window(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)

        # ====== Functionality code starts here ======
        self.pushButton_4.clicked.connect(self.openfile)
        self.pushButton_3.clicked.connect(self.datahead)

        self.connectButtons()
        self.statusBar.showMessage(">")

    def connectButtons(self):
        self.actionExit.triggered.connect(self.close)
    
    def fill_table_json(self):
        with open(self.main_table_path, 'r') as table_file_read:
            self.data = json.load(table_file_read)
        
        the_keys = list(self.data['document_list_and_data'][0].keys())
        the_list = self.data['document_list_and_data']

        self.tableWidget.setColumnCount(len(the_keys))
        self.tableWidget.setHorizontalHeaderLabels(the_keys)
        self.tableWidget.setRowCount(len(the_list))

        print(f"The type is: {type(the_keys)}, and the content is: {the_keys}")

        for i in range(len(the_list)): #row count
            for j in range(len(the_keys)): #column count
                cell_val = the_list[i][the_keys[j]]
                self.tableWidget.setItem(i, j, QTableWidgetItem(cell_val))

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

    def openfile(self):
        try:
            # path = QFileDialog.getOpenFileName(self, 'Open JSON file', os.getenv('HOME'), 'JSON(*.json)')[0]
            #self.all_data = json.load(path)

            path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/listing.json"
            with open(path, 'r') as file:
                self.data = json.load(file)
        except:
            print("Unexpected error 1:", sys.exc_info()[0])

        try:
            # print(f"\nThe file type is: {type(self.data)}")
            # print(f"\n{self.data}")

            # print(f"\nThe first key is: {self.data['document_list_and_data'][0].keys()}")

            # for entrykey in self.data['document_list_and_data']:
            #     print(entrykey.keys())

            # for entry in self.data['document_list_and_data']:
            #     print(entry['ID'])
            the_keys = list(self.data['document_list_and_data'][0].keys())
            the_list = self.data['document_list_and_data']
        except:
            print("\nUnexpected error 2:", sys.exc_info()[0])

        try:
            self.tableWidget.setColumnCount(len(the_keys))
            self.tableWidget.setHorizontalHeaderLabels(the_keys)
            self.tableWidget.setRowCount(len(the_list))
            
        except:
            print("Unexpected error 3:", sys.exc_info())

        try:
            for i in range(len(the_list)): #row count
                for j in range(len(the_keys)): #column count
                    cell_val = the_list[i][the_keys[j]]
                    self.tableWidget.setItem(i, j, QTableWidgetItem(cell_val))
                    
        except:
            print("Unexpected error 4:", sys.exc_info()[0])

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()


    def datahead(self):
        num_column = 2
        if num_column == 0:
            num_rows = len(self.data.index)
        else:
            num_rows = num_column

        try:
            self.tableWidget.setColumnCount(len(self.data.columns))
            self.tableWidget.setRowCount(num_rows)
            self.tableWidget.setHorizontalHeaderLabels(self.data.columns)
        except:
            print("Unexpected error 3:", sys.exc_info())

        try:
            for i in range(num_rows):
                for j in range(len(self.data.columns)):
                    self.tableWidget.setItem(
                        i, j, QTableWidgetItem(str(self.data.iat[i, j])))
        except:
            print("Unexpected error 4:", sys.exc_info()[0])

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

    # def scanDirectoriesForFiles(self):
    #     try:
            
    #     except:
    #         print("Unexpected error 3:", sys.exc_info())


# class FindReplaceDialog(QDialog):
#     def __init__(self, parent=None):
#         super().__init__(parent)
#         loadUi("ui/find_replace.ui", self)
#         self.pushButton.clicked.connect(self.addtesttomainbox)
#
#     def addtesttomainbox(self):
#         win.textEdit.append("Hello World")

# ====== Functionality code ends here ======


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.show()
    sys.exit(app.exec())
