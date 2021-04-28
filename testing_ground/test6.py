import sys
import os
from os.path import dirname, realpath, join
# from PyQt5.QtWidgets import QDialogButtonBox, QApplication, QDialog, QWidget, QFileDialog, QTableWidget, QTableWidgetItem, QMainWindow, QLineEdit, QSpinBox
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import json
from PyQt5.uic import loadUi, loadUiType
import sqlite3

From_Main, _ = loadUiType(join(dirname(__file__), "app2.ui"))


class MainWindow(QMainWindow, From_Main):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        # ====== Functionality code starts here ======
        self.statusBar.showMessage(">")

        self.pushButton_4.clicked.connect(self.create_task_qualification_report)
        self.pushButton_3.clicked.connect(self.create_new_sop)
        self.pushButton_2.clicked.connect(self.create_procedure_checklist)

        self.connectButtons()
        self.load_global_variables()
        # self.fill_table_json()
        self.fill_table_sqlite()

    def connectButtons(self):
        self.actionSystem_parameters.triggered.connect(self.open_pth_edit_dialogue)
        self.actionExit.triggered.connect(self.close)

    def load_global_variables(self):
        path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/global_var_2.json"
        with open(path, 'r') as config_file_read:
            self.loaded_vars = json.load(config_file_read)

        # print(loaded_vars[list(loaded_vars.keys())[0]])
        self.sop_path = self.loaded_vars["sop_doc_folder"]
        self.tqr_path = self.loaded_vars["task_qual_rep_folder"]
        self.pc_path = self.loaded_vars["proc_checklist_folder"]
        self.temp_path = self.loaded_vars["template_doc_folder"]
        self.main_table_path = self.loaded_vars["main_table_data_file"]
        self.file_expiry = self.loaded_vars["file_expiry_time_months"]

        # self.statusBar.showMessage(f"Loaded vars: {type(self.loaded_vars)}")

    def fill_table_sqlite(self):
        # main_table_path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/test.db"
        connection = sqlite3.connect(self.main_table_path)
        query = "SELECT * FROM main_table_db"
        result = connection.execute(query)

        query2 = "PRAGMA table_info(main_table_db)"
        result2 = connection.execute(query2)

        header_list = result2.fetchall()
        header_list_2 = []

        for i in header_list:
            header_list_2.append(i[1])

        self.tableWidget.setColumnCount(len(header_list_2))
        self.tableWidget.setHorizontalHeaderLabels(list(header_list_2))

        for row_number, row_data in enumerate(result):
            self.tableWidget.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                # print(f"1 = Row: {row_number}, Col: {column_number}, Data: {data}")
                if (column_number == 4) and (data == 0):
                    self.tableWidget.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                    # QTableWidgetItem
                    self.tableWidget.item(row_number, column_number).setTextAlignment(Qt.AlignCenter)
                    self.tableWidget.item(row_number, column_number).setBackground(QColor(255,0,0))
                    # print(f"2 = Row: {row_number}, Col: {column_number}, Data: {data}")
                else:
                    self.tableWidget.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                    # self.tableWidget.item(row_number, column_number).setBackground(QColor(255,255,255))
                    # print(f"3 = Row: {row_number}, Col: {column_number}, Data: {data}")

        connection.close()

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

    def create_task_qualification_report(self):
        None

    def create_new_sop(self):
        dialogue1 = NewSOPDialogue(self)
        dialogue1.exec()

    def create_procedure_checklist(self):
        None

    def rescan_dir_for_json_file(self):
        # SELECT * FROM main_table_db WHERE expiry_status glob '**';
        None

    def open_pth_edit_dialogue(self):
        dialogue1 = FileFolderPathDialogue(self)
        dialogue1.exec()


class FileFolderPathDialogue(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi("C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/doc_dir.ui", self)
        
        # Initialise OK-Apply-Cancel button functionality
        self.pushButton.clicked.connect(self.save_close_changes)
        self.pushButton_2.clicked.connect(self.save_changes)
        self.pushButton_3.clicked.connect(self.close)

        # Initialise toolbutton functionality
        self.toolButton.clicked.connect(self.choose_path_sops)
        self.toolButton_2.clicked.connect(self.choose_path_task_qual_report)
        self.toolButton_3.clicked.connect(self.choose_path_procedure_checklist)
        self.toolButton_4.clicked.connect(self.choose_path_template_folder)
        self.toolButton_5.clicked.connect(self.choose_path_main_table)

        # Populate field values/data
        self.lineEdit.setText(mywindow.sop_path)
        self.lineEdit_2.setText(mywindow.tqr_path)
        self.lineEdit_3.setText(mywindow.pc_path)
        self.lineEdit_4.setText(mywindow.temp_path)
        self.lineEdit_5.setText(mywindow.main_table_path)
        self.spinBox.setValue(mywindow.file_expiry)

    def save_changes(self):

        global_variable_dict = {
            "sop_doc_folder": self.lineEdit.text(),
   	        "task_qual_rep_folder": self.lineEdit_2.text(),
   	        "proc_checklist_folder": self.lineEdit_3.text(),
   	        "template_doc_folder": self.lineEdit_4.text(),
   	        "main_table_data_file": self.lineEdit_5.text(),
   	        "file_expiry_time_months": self.spinBox.value()
        }

        test_globvar = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/global_var_2.json"
        config_file_write = open(test_globvar, 'w')
        json.dump(global_variable_dict, config_file_write)

        mywindow.sop_path = self.lineEdit.text()
        mywindow.tqr_path = self.lineEdit_2.text()
        mywindow.pc_path = self.lineEdit_3.text()
        mywindow.temp_path = self.lineEdit_4.text()
        mywindow.main_table_path = self.lineEdit_5.text()
        mywindow.file_expiry = self.spinBox.value()

        mywindow.statusBar.showMessage("Apply: Paths sucessfully updated")

    def save_close_changes(self):
        self.save_changes()
        mywindow.statusBar.showMessage("OK: Paths sucessfully updated")
        self.close()

    def choose_path_sops(self):
        path = QFileDialog.getExistingDirectory(self, 'Select folder containing SOP documents', os.getenv('HOME'))
        if path != "":
            self.lineEdit.setText(path)

    def choose_path_task_qual_report(self):
        path = QFileDialog.getExistingDirectory(self, 'Select folder containing task qualification report documents', os.getenv('HOME'))
        if path != "":
            self.lineEdit_2.setText(path)

    def choose_path_procedure_checklist(self):
        path = QFileDialog.getExistingDirectory(self, 'Select folder containing procedure checklist documents', os.getenv('HOME'))
        if path != "":
            self.lineEdit_3.setText(path)

    def choose_path_template_folder(self):
        path = QFileDialog.getExistingDirectory(self, 'Select folder containing template documents', os.getenv('HOME'))
        if path != "":
            self.lineEdit_4.setText(path)

    def choose_path_main_table(self):
        path = QFileDialog.getOpenFileName(self, 'Select file for populating main table', os.getenv('HOME'), 'SQLite Database(*.db)')[0]
        if path != "":
            self.lineEdit_5.setText(path)


class NewSOPDialogue(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi("C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/new_sop.ui", self)

        # Initialise OK-Cancel button functionality

        # self.buttonBox.button(QDialogButtonBox.Ok).clicked.connect(lambda: self.save_changes("ok"))
        # self.buttonBox.button(QDialogButtonBox.Cancel).clicked.connect(lambda: self.save_changes("cancel"))

        self.buttonBox.button(QDialogButtonBox.Ok).clicked.connect(self.save_changes)

    def save_changes(self):
        name = self.lineEdit.text()

        mywindow.statusBar.showMessage(f"New SOP: {name}, has been created")

# ====== Functionality code ends here ======


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = MainWindow()
    mywindow.show()
    sys.exit(app.exec_())
