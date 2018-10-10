import sys
import os
import platform
from pathlib import (PureWindowsPath, Path)
from PyQt5.QtWidgets import (QApplication, QMainWindow)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import *
import xlsxwriter

# License class constructor
class License(object):

    #define and validate License variables
    def __init__(self, site, type, numcam, numint, numLPR, existing, fdob):

        # verify site is global office or dc/pop/colo, if is, assign to self.site
        # if "global office" not in site.lower() and "dc/pop/colo" not in site.lower():
        #     # if not, raise exception
        #     raise ValueError
        self.site = site

        # verify type is new site or expansion, if not, raise exception
        # if is, assign to self.type
        # if "new site" not in type.lower() and "expansion" not in type.lower():
        #     raise ValueError
        self.type = type
        # verify that SOMETHING is in numcam, if not just make it zero
        # if numcam == "":
        #     numcam = 0
        # assign to numcam and make integer
        self.numcam = int(numcam)
        # verify that SOMETHING is in numint, if not, just make it zero
        # if numint == "":
        #     numint = 0
        # assign to numint and make integer
        self.numint = int(numint)
        # verify that SOMETHING is in numLPR, if not, just make it zero
        # if numLPR == "":
        #     numLPR = 0
        # assign to numLPR and make integer
        self.numLPR = int(numLPR)
        # verify that existing is yes or no, if not, raise exception
        # if "yes" not in existing.lower() and "no" not in existing.lower():
        #     raise ValueError
        self.existing = existing
        # test = fdob.replace(", ", "")
        # if len(test) != 8:
        #     raise ValueError
        self.fdob = fdob

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(681, 495)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(350, 30, 161, 32))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(130, 36, 211, 20))
        self.label.setObjectName("label")
        self.comboBox_2 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_2.setGeometry(QtCore.QRect(350, 60, 161, 32))
        self.comboBox_2.setObjectName("comboBox_2")
        self.comboBox_2.addItem("")
        self.comboBox_2.addItem("")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(180, 66, 161, 20))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(90, 130, 251, 20))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(190, 160, 151, 20))
        self.label_4.setObjectName("label_4")
        self.comboBox_5 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_5.setGeometry(QtCore.QRect(350, 90, 161, 32))
        self.comboBox_5.setObjectName("comboBox_5")
        self.comboBox_5.addItem("")
        self.comboBox_5.addItem("")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(170, 90, 171, 31))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(140, 185, 201, 31))
        self.label_6.setObjectName("label_6")
        self.dateEdit = QtWidgets.QDateEdit(self.centralwidget)
        self.dateEdit.setGeometry(QtCore.QRect(360, 230, 151, 22))
        self.dateEdit.setDateTime(QtCore.QDateTime(QtCore.QDate(2018, 10, 24), QtCore.QTime(0, 0, 0)))
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setObjectName("dateEdit")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(290, 230, 61, 20))
        self.label_7.setObjectName("label_7")
        self.spinBox = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox.setGeometry(QtCore.QRect(360, 130, 141, 22))
        self.spinBox.setObjectName("spinBox")
        self.spinBox_2 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_2.setGeometry(QtCore.QRect(360, 160, 141, 22))
        self.spinBox_2.setObjectName("spinBox_2")
        self.spinBox_3 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_3.setGeometry(QtCore.QRect(360, 190, 141, 22))
        self.spinBox_3.setObjectName("spinBox_3")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(400, 280, 114, 32))
        self.pushButton.setAutoDefault(False)
        self.pushButton.setDefault(False)
        self.pushButton.setObjectName("pushButton")
        self.lic = self.pushButton.clicked.connect(self.calculate)
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(230, 280, 114, 32))
        self.pushButton_2.setObjectName("pushButton_2")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 681, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)







    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Genetec License Calculator"))
        self.comboBox.setItemText(0, _translate("MainWindow", "Global Office"))
        self.comboBox.setItemText(1, _translate("MainWindow", "DC/PoP/COLO"))
        self.label.setText(_translate("MainWindow", "Global Office or DC/PoP/COLO?"))
        self.comboBox_2.setItemText(0, _translate("MainWindow", "New Site"))
        self.comboBox_2.setItemText(1, _translate("MainWindow", "Expansion"))
        self.label_2.setText(_translate("MainWindow", "New Site or Expansion?"))
        self.label_3.setText(_translate("MainWindow", "Number of Cameras, Excluding LPR\'s?"))
        self.label_4.setText(_translate("MainWindow", "Number of Intercoms?"))
        self.comboBox_5.setItemText(0, _translate("MainWindow", "No"))
        self.comboBox_5.setItemText(1, _translate("MainWindow", "Yes"))
        self.label_5.setText(_translate("MainWindow", "Existing LPRs in System?"))
        self.label_6.setText(_translate("MainWindow", "Number of LPR\'s to be Added?"))
        self.dateEdit.setDisplayFormat(_translate("MainWindow", "MM/dd/yyyy"))
        self.label_7.setText(_translate("MainWindow", "FDOB?"))
        self.pushButton.setText(_translate("MainWindow", "Calculate"))
        self.pushButton_2.setText(_translate("MainWindow", "Cancel"))

    def calculate(self):
        self.site = self.comboBox.currentText()
        self.type = self.comboBox_2.currentText()
        self.numcam = int(self.spinBox.value())
        self.numint = int(self.spinBox_2.value())
        self.numLPR = int(self.spinBox_3.value())
        self.existing = self.comboBox_5.currentText()
        self.fdob = self.dateEdit.date().toPyDate()

        self.lic = License(self.site, self.type, self.numcam, self.numint,
                                    self.numLPR, self.existing, self.fdob)

        self.system = [self.site, self.type, self.existing, self.numcam,
                        self.numint, self.numLPR]

        self.licenses = []
        self.qty = []



        #datetime
        self.today = date.today()

        #end of current year
        self.eoy = [self.today.year, 12, 31]

        #beginning of desired SMA year
        self.boyear = [(self.today.year + 2)]

        #format today
        self.today = [self.today.year, self.today.month, self.today.day]

        #diff in months from eoy to today and boyear
        self.diff = (self.eoy[1] - self.lic.fdob.month)
        self.diff_y = (self.boyear[0] - self.lic.fdob.year)

        if self.lic.type.lower() == "new site":
            self.licenses.append("GSC-Base-5.7")
            self.qty.append(1)
            self.licenses.append("GSC-1AD-USCH")
            self.qty.append(1)
            self.licenses.append("GSC-Om-E")
            self.qty.append(1)
            self.licenses.append("GSC-1SDK-SUREVIEW-Immix")
            self.qty.append(1)


        if self.lic.site.lower() == "dc/pop/colo" and self.lic.type.lower() == "new site":
            self.licenses.append("GSC-PM-STD-SiteLicense")
            self.qty.append(1)
            self.licenses.append("GSC-1U")
            self.qty.append(15)
            self.licenses.append("GSC-1FOD")
            self.qty.append(1)

        if self.lic.site.lower() == "dc/pop/colo" or self.lic.site.lower() == "global office":
            self.licenses.append("GSC-Om-E-1C")
            self.qty.append(self.lic.numcam + self.lic.numint)
            self.licenses.append("ADV-CAM-E-1M")
            self.qty.append(self.diff * (self.lic.numcam + self.lic.numint))
            self.licenses.append("ADV-CAM-E-1Y")
            self.qty.append(self.diff_y * (self.lic.numcam + self.lic.numint))

        if self.lic.site.lower() == "dc/pop/colo":
            self.licenses.append("GSC-OM-E-1FC")
            self.qty.append(self.lic.numcam + self.lic.numint)

        if self.lic.existing.lower() == "no" and self.lic.numLPR != 0:
            self.licenses.append("GSC-Av-S")
            self.qty.append(1)
        elif self.lic.numLPR !=0:
            self.licenses.append("GSC-Av-S-1SHP")
            self.qty.append(self.lic.numLPR)
            self.licenses.append("ADV-LPR-F-1M")
            self.qty.append(self.diff * self.lic.numLPR)
            self.licenses.append("ADV-LPR-F-1Y")
            self.qty.append(self.diff_y * self.lic.numLPR)


        if platform.system() == 'Windows':
            fn = str(PureWindowsPath('/Genetec_Licenses.xlsx'))
            dtop = str(PureWindowsPath(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')))
            pth = (dtop + fn)
        else:
            dtop = os.path.join(os.path.join(os.path.expanduser('~')))
            fn = '/Desktop/Genetec_Licenses.xlsx'
            pth = (dtop + fn)
        wb = xlsxwriter.Workbook(pth)
        ws = wb.add_worksheet()
        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 12)

        bold = wb.add_format({'bold': True})
        cf = wb.add_format()
        cf.set_align('left')
        ul = wb.add_format({'underline': True, 'bold': True})
        d_format = wb.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})

        ws.write('A1', 'Your System', ul)
        ws.write('A2', 'Site:')
        ws.write('A3', 'Type: ')
        ws.write('A4', 'Existing LPRs on Site?: ')
        ws.write('A5', 'Number of Cameras: ')
        ws.write('A6', 'Number of Intercoms: ')
        ws.write('A7', 'Number of LPRs: ')
        ws.write('A8', 'FDOB: ')
        ws.write('A10', 'Genetec License Part No.', bold)
        ws.write('B10', 'Quantity', bold)

        row = 1
        col = 1
        for s in (self.system):
            ws.write(row, col,   s, cf)
            row += 1
        ws.write(row, col, self.fdob, d_format)

        row = 10
        col = 0

        for lic in (self.licenses):
            ws.write(row, col,   lic, cf)
            row += 1
        row = 10
        for quantity in (self.qty):
            ws.write(row, col + 1, quantity, cf)
            row += 1

        wb.close()

        w.close()

app = QApplication(sys.argv)
w = QMainWindow()
ui = Ui_MainWindow()
y = ui.setupUi(w)

w.show()
sys.exit(app.exec_())
