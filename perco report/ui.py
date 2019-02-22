# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mainwindow.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(460, 360)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btn_report = QtWidgets.QPushButton(self.centralwidget)
        self.btn_report.setGeometry(QtCore.QRect(10, 10, 150, 20))
        self.btn_report.setObjectName("btn_report")
        self.lst = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.lst.setGeometry(QtCore.QRect(200, 90, 250, 200))
        self.lst.setObjectName("lst")
        self.btn_rt = QtWidgets.QPushButton(self.centralwidget)
        self.btn_rt.setGeometry(QtCore.QRect(10, 40, 150, 20))
        self.btn_rt.setObjectName("btn_rt")

        self.btn_add_staff = QtWidgets.QPushButton(self.centralwidget)
        self.btn_add_staff.setGeometry(QtCore.QRect(10, 90, 150, 20))
        self.btn_add_staff.setObjectName("btn_add_staff")

        self.dateStart = QtWidgets.QDateEdit(self.centralwidget)
        self.dateStart.setGeometry(QtCore.QRect(200, 40, 110, 22))
        self.dateStart.setObjectName("dateStart")
        self.dateEnd = QtWidgets.QDateEdit(self.centralwidget)
        self.dateEnd.setGeometry(QtCore.QRect(330, 40, 110, 22))
        self.dateEnd.setObjectName("dateEnd")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(200, 10, 250, 22))
        self.comboBox.setObjectName("comboBox")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 455, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.btn_report.setText(_translate("MainWindow", "Отчёт"))
        self.btn_rt.setText(_translate("MainWindow", "Табель из карты"))
        self.btn_add_staff.setText(_translate("MainWindow", "Добавить Сотрудников"))
