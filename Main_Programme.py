# General requirements
import requests
import json
import secrets
import datetime
import os
import sys

# QT5 elements
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMainWindow
from PyQt5.QtGui import QIcon
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import QUrl

# Word doc elements
from docx import Document
from PIL import Image, ImageOps

# Excel doc elements
from openpyxl import load_workbook

# Logging elements
from exception_decor import exception
from exception_logger import logger

# Airtable elements
airtable_api_key = secrets.airtable_api_key
base_key = secrets.base_key
table_name = secrets.table_name

from MainUI import Ui_MainWindow

## Some bits for Pyinstaller to help it find relative path files ##
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))


@exception(logger)
class MainWindow:
    def __init__(self):
        self.main_win = QMainWindow()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self.main_win)
        QtCore.QCoreApplication.processEvents()

        self.ui.pushButton_generate_reports.clicked.connect(self.generate_reports)

    def show(self):
        self.main_win.show()
        self.ui.listWidget_outputWindow.clear()
        self.ui.listWidget_outputWindow.addItem("Welcome to the Site Visit Report tool")
        self.ui.listWidget_outputWindow.addItem("Enter the required information then click 'Generate Reports'")
        self.ui.listWidget_outputWindow.addItem("All items must be filled in!")


    def generate_reports(self):
        svr_no = self.ui.lineEdit_svr_no.text()
        site_visit_date = self.ui.lineEdit_site_visit_date.text()
        date1 = self.ui.lineEdit_date1.text()
        date2 = self.ui.lineEdit_date2.text()
        surveyor_names = self.ui.lineEdit_present_for_survey.text()
        chaperone = self.ui.lineEdit_chaperone.text()
        issued_by = self.ui.lineEdit_issued_by.text()
        progress_from_site = self.ui.textEdit_progress_notes.toPlainText()

        if svr_no == "":
            self.ui.listWidget_outputWindow.addItem("SVR number empty - please add value")
            return 0
        if site_visit_date == "":
            self.ui.listWidget_outputWindow.addItem("Site Visit Date empty - please add value")
            return 0
        if date1 == "":
            self.ui.listWidget_outputWindow.addItem("Start date empty - please add value")
            return 0
        if date2 == "":
            self.ui.listWidget_outputWindow.addItem("End date empty - please add value")
            return 0
        if surveyor_names == "":
            self.ui.listWidget_outputWindow.addItem("Surveyor Names empty - please add value")
            return 0
        if chaperone == "":
            self.ui.listWidget_outputWindow.addItem("Chaperone name empty - please add value")
            return 0
        if issued_by == "":
            self.ui.listWidget_outputWindow.addItem("Issued by field empty - please add value")
            return 0
        if progress_from_site == "":
            self.ui.listWidget_outputWindow.addItem("Please add progress notes from site")
            return 0
        else:
            self.ui.listWidget_outputWindow.clear()
            self.ui.listWidget_outputWindow.addItem("Carrying on")








if __name__ == '__main__':
    QtWidgets.QApplication.setAttribute(QtCore.Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())
