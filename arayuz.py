# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'arayuz.ui'
#
# Created by: PyQt5 UI code generator 5.14.1
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 525)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.load_lists_button = QtWidgets.QPushButton(self.centralwidget)
        self.load_lists_button.setGeometry(QtCore.QRect(30, 30, 140, 32))
        self.load_lists_button.setObjectName("load_lists_button")
        self.selected_file_text = QtWidgets.QLabel(self.centralwidget)
        self.selected_file_text.setGeometry(QtCore.QRect(190, 30, 300, 32))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.selected_file_text.setFont(font)
        self.selected_file_text.setText("")
        self.selected_file_text.setObjectName("selected_file_text")
        self.create_list_button = QtWidgets.QPushButton(self.centralwidget)
        self.create_list_button.setGeometry(QtCore.QRect(30, 140, 150, 28))
        self.create_list_button.setObjectName("create_list_button")
        self.list_widget = QtWidgets.QListWidget(self.centralwidget)
        self.list_widget.setGeometry(QtCore.QRect(500, 70, 256, 421))
        self.list_widget.setObjectName("list_widget")
        self.okul_adi_text_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.okul_adi_text_2.setGeometry(QtCore.QRect(100, 90, 251, 22))
        self.okul_adi_text_2.setObjectName("okul_adi_text_2")
        self.okul_adi_text = QtWidgets.QLabel(self.centralwidget)
        self.okul_adi_text.setGeometry(QtCore.QRect(30, 90, 55, 22))
        self.okul_adi_text.setObjectName("okul_adi_text")
        self.save_list_button = QtWidgets.QPushButton(self.centralwidget)
        self.save_list_button.setGeometry(QtCore.QRect(30, 200, 150, 28))
        self.save_list_button.setObjectName("save_list_button")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(500, 20, 251, 22))
        self.label.setObjectName("label")
        MainWindow.setCentralWidget(self.centralwidget)

        self.load_lists_button.clicked.connect(self.yukle)
        self.create_list_button.clicked.connect(self.olustur)
        self.save_list_button.clicked.connect(self.kaydet)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        # QFileDialog.getOpenFileName("")

    def yukle(self):
        file_name, _ = QFileDialog.getOpenFileName(filter="Excell Dosyası (*.xlsx)")
        if file_name:
            print(file_name)
            self.selected_file_text.setText(file_name)

    def olustur(self):
        okul = self.okul_adi_text_2.text()
        if okul:
            print("Olusturuldu")
        self.ekle()
    def kaydet(self):
        okul = self.okul_adi_text_2.text()
        if okul:
            print("Kaydedildi")

    def ekle(self):
        self.list_widget.clear()
        self.list_widget.addItem("NU")

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.load_lists_button.setText(_translate("MainWindow", "Öğrenci Listesi Yükle"))
        self.create_list_button.setText(_translate("MainWindow", "Yoklamaları Oluştur"))
        self.okul_adi_text.setText(_translate("MainWindow", "Okul Adı:"))
        self.save_list_button.setText(_translate("MainWindow", "Yoklamaları Kaydet"))
        self.label.setText(_translate("MainWindow", "Oluşturulacak Yoklama Listeleri"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
