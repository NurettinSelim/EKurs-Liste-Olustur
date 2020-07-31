import sys

from PyQt5 import QtWidgets, QtGui, uic
from PyQt5.QtWidgets import QFileDialog

from course_list_generator import create_list, save_list_to_excel


def show_message(message):
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setText(message)
    msg.setWindowTitle("Uyarı")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('arayuz.ui', self)
        self.setWindowIcon(QtGui.QIcon('icon.png'))

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

        # - Variables - #
        self.file_name = str()
        self.course_data = dict()

        self.show()

    def select_list(self):
        self.file_name, _ = QFileDialog.getOpenFileName(filter="Excell Dosyası (*.xlsx)")
        if self.file_name:
            if len(self.file_name) > 40:
                self.file_name_label.setText(self.file_name[self.file_name.find("/", 3):])
            else:
                self.file_name_label.setText(self.file_name)

    def create_list(self):
        if self.file_name:
            try:
                self.course_data = create_list(self.file_name)
                self.create_list_view(self.course_data["lists"])
            except Exception as e:
                print(e)
                show_message("Hatalı liste formatı!")
        else:
            show_message("Lütfen tercih listesini seçiniz!")

    def save_list(self):
        schoolname = self.okul_adi_line_edit.text()
        if schoolname:
            if self.file_name:
                if self.course_data:
                    filename, _ = QFileDialog.getSaveFileName(filter="Excell Dosyası (*.xlsx)")
                    if filename:
                        save_list_to_excel(filename=filename,
                                           schoolname=schoolname,
                                           data_dict=self.course_data)
                else:
                    show_message("Lütfen yoklamaları oluşturunuz!")
            else:
                show_message("Lütfen tercih listesini seçiniz!")
        else:
            show_message("Lütfen okul adı giriniz!")

    def create_list_view(self, lists):
        upper_lists = map(lambda x: x.upper(), lists)
        self.list_widget.clear()
        self.list_widget.addItems(upper_lists)


app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()
