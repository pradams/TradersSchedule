from PyQt5.QtWidgets import QLabel, QWidget, QPushButton, QGridLayout, QComboBox, QFileDialog, QLineEdit, QDialog
from PyQt5.QtCore import QRect, QMetaObject, QCoreApplication
from ExcelWriter import ExcelWriter
from VisualElement import VisualTable
from xlutils.copy import copy
from xlutils.styles import Styles
import xlrd

class MainIntroWindow(QDialog):

    def __init__(self):
        super().__init__()
        self.original_filename = ''
        self.new_save_filename = ''
        self.default_open_path = "/Users/Patrick/Desktop/TradersSchedule/test.xls"
        self.default_save_path = "/Users/Patrick/Desktop/TradersSchedule/new_schedule.xls"
        self.open_file_found = False
        self.save_file_found = False
        self.day_index = 0
        self.setupUi(self)

    def setupUi(self, Dialog):
        self.setWindowTitle("Trader's Scheduler")
        Dialog.resize(574, 270)
        self.day_selector = QComboBox(Dialog)
        self.day_selector.setGeometry(QRect(250, 80, 121, 26))
        self.day_selector.setEditable(False)
        self.day_selector.setCurrentText("")
        self.day_selector.setObjectName("day_selector")
        self.day_selector.addItems(["Monday", "Tuesday", "Wednesday", "Thursday",
                                    "Friday", "Saturday", "Sunday", "All"])
        self.label = QLabel(Dialog)
        self.label.setGeometry(QRect(170, 80, 81, 20))
        self.label.setObjectName("label")
        self.open_file_button = QPushButton(Dialog)
        self.open_file_button.setGeometry(QRect(30, 120, 191, 31))
        self.open_file_button.setParent(Dialog)
        self.open_file_button.setObjectName("open_file_button")
        self.open_file_button.clicked.connect(self.handleOpenFileButtonClick)
        self.lineEdit = QLineEdit(Dialog)
        self.lineEdit.setGeometry(QRect(220, 125, 300, 21))
        self.lineEdit.setReadOnly(True)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QLineEdit(Dialog)
        self.lineEdit_2.setGeometry(QRect(220, 155, 300, 21))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.set_save_file_button = QPushButton(Dialog)
        self.set_save_file_button.setGeometry(QRect(50, 150, 171, 31))
        self.set_save_file_button.setObjectName("set_save_file_button")
        self.set_save_file_button.clicked.connect(self.handleSaveFileButtonClick)
        self.run_program_button = QPushButton(Dialog)
        self.run_program_button.setGeometry(QRect(220, 190, 144, 32))
        self.run_program_button.setObjectName("run_program_button")
        self.run_program_button.clicked.connect(self.runProgram)
        self.run_program_button.setDisabled(True)

        # Set up default file parameters
        self.original_filename = self.default_open_path
        self.lineEdit.setText(self.original_filename)
        if self.original_filename:
            self.open_file_found = True
        self.new_save_filename = self.default_save_path
        self.lineEdit_2.setText(self.new_save_filename)
        if self.new_save_filename:
            self.save_file_found = True

        if self.open_file_found and self.save_file_found:
            self.run_program_button.setDisabled(False)
        else:
            self.run_program_button.setDisabled(True)

        self.retranslateUi(Dialog)
        QMetaObject.connectSlotsByName(Dialog)
        self.show()

    def retranslateUi(self, Dialog):
        _translate = QCoreApplication.translate
        self.label.setText(_translate("Dialog", "Select a Day:"))
        self.open_file_button.setText(_translate("Dialog", "Choose Original Filename:"))
        self.set_save_file_button.setText(_translate("Dialog", "Choose New Filename:"))
        self.run_program_button.setText(_translate("Dialog", "Set Up Schedule"))

    def runProgram(self):
        self.day_index = self.day_selector.currentIndex()

        # Setup and copy workbook.
        book = xlrd.open_workbook(filename=self.original_filename, formatting_info=True, on_demand=True)
        new_schedule = copy(book)

        excelWriter = ExcelWriter(book, new_schedule.get_sheet(self.day_index + 3),
                                  book.sheet_by_index(self.day_index + 3),
                                  new_schedule,
                                  Styles(book), self.new_save_filename)

        # Set break and lunch times in new excel file.
        excelWriter.calcBreakTimes()
        excelWriter.calcLunchTimes()

        # Set up schedule (Color Cells)
        excelWriter.colorCells()

        # Open editing table.
        ex = VisualTable(self.new_save_filename, excelWriter.calcNumEmployees(), self.day_index)
        ex.createTable()

        # Close opening main window.
        self.close()

    def handleOpenFileButtonClick(self):
        file_dialog = QFileDialog()
        file_dialog.move(300, 300)
        options = file_dialog.Options()
        options |= file_dialog.DontUseNativeDialog

        self.original_filename, _ = file_dialog.getOpenFileName(self, "Choose a schedule", self.default_open_path, "Excel Files (*.xls)", options=options)
        if self.original_filename:
            self.open_file_found = True
            self.lineEdit.setText(self.original_filename)
        else:
            self.open_file_found = False

        if self.open_file_found and self.save_file_found:
            self.run_program_button.setDisabled(False)
        else:
            self.run_program_button.setDisabled(True)

    def handleSaveFileButtonClick(self):
        file_dialog = QFileDialog()
        file_dialog.move(300,300)
        options = file_dialog.Options()
        options |= file_dialog.DontUseNativeDialog

        self.new_save_filename, _ = file_dialog.getSaveFileName(self, "Choose new filename", self.default_save_path, "Excel Files (*.xls)", options= options)
        if self.new_save_filename[-4:] != ".xls":
            self.new_save_filename = self.new_save_filename + ".xls"

        if self.new_save_filename != ".xls":
            self.save_file_found = True
            self.lineEdit_2.setText(self.new_save_filename)
        else:
            self.save_file_found = False

        if self.open_file_found and self.save_file_found:
            self.run_program_button.setDisabled(False)
        else:
            self.run_program_button.setDisabled(True)

    def handleDayIndexChanged(self, index):
        self.day_index = index