import sys
import os
from os.path import dirname, realpath, join
from PyQt5.QtWidgets import QApplication, QDialog, QWidget, QFileDialog, QTableWidget, QTableWidgetItem, QMainWindow
import json
from PyQt5.uic import loadUi, loadUiType

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
        self.fill_table()

    def connectButtons(self):
        self.actionSystem_parameters.triggered.connect(self.open_pth_edit_dialogue)
        self.actionExit.triggered.connect(self.close)

    def load_global_variables(self):
        path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/global_var.json"
        with open(path, 'r') as file:
            self.loaded_vars = json.load(file)

    def fill_table(self):
        path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/listing.json"
        with open(path, 'r') as file:
            self.data = json.load(file)
        
        the_keys = list(self.data['document_list_and_data'][0].keys())
        the_list = self.data['document_list_and_data']

        self.tableWidget.setColumnCount(len(the_keys))
        self.tableWidget.setHorizontalHeaderLabels(the_keys)
        self.tableWidget.setRowCount(len(the_list))

        for i in range(len(the_list)): #row count
            for j in range(len(the_keys)): #column count
                cell_val = the_list[i][the_keys[j]]
                self.tableWidget.setItem(i, j, QTableWidgetItem(cell_val))

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()

    def create_task_qualification_report(self):
        None

    def create_new_sop(self):
        None

    def create_procedure_checklist(self):
        None
    
    def rescan_dir_for_json_file(self):
        sop_path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/SOP_docs/sops"
        tqr_path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/SOP_docs/task_qual_reports"
        pc_path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/SOP_docs/procedure_checklists"

    def open_pth_edit_dialogue(self):
        dialogue1 = FileFolderPathDialogue(self)
        dialogue1.exec()


class FileFolderPathDialogue(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi("testing_ground/doc_dir.ui", self)
        self.pushButton.clicked.connect(self.save_changes)

    def save_changes(self):
        mywindow.statusBar.showMessage("Paths sucessfully updated")
# ====== Functionality code ends here ======


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = MainWindow()
    mywindow.show()
    sys.exit(app.exec_())
