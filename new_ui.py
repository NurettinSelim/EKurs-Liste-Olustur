import sys

from PyQt5 import QtWidgets, uic


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('arayuz.ui', self)

        # - Defining Items - #
        # Buttons
        self.select_list_button = self.findChild(QtWidgets.QPushButton, "selectListButton")
        self.create_list_button = self.findChild(QtWidgets.QPushButton, "createListButton")
        self.save_list_button = self.findChild(QtWidgets.QPushButton, "saveListButton")

        # Labels
        self.file_name_label = self.findChild(QtWidgets.QLabel, "FileNameLabel")

        # LineEdits
        self.okul_adi_line_edit = self.findChild(QtWidgets.QLineEdit, "okulAdiLineEdit")

        # ListWidgets
        self.list_widget = self.findChild(QtWidgets.QListWidget, "listWidget")

        # - Events - #
        # Button Events
        self.select_list_button.clicked.connect(self.select_list)
        self.create_list_button.clicked.connect(self.create_list)
        self.save_list_button.clicked.connect(self.save_list)

        self.show()

    def select_list(self):
        print("select")

    def create_list(self):
        print("create")

    def save_list(self):
        print("save")


app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()
