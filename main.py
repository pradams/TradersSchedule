import xlrd
import sys
from VisualElement import VisualTable
from VisualElement import MainIntroWindow
from ExcelWriter import ExcelWriter
from xlutils.copy import copy
from xlutils.styles import Styles
from PyQt5.QtWidgets import QApplication, QFileDialog

### Create Applicatin Window ###
app = QApplication(sys.argv)

# Create filebrowser and day selector.
#browser = MainIntroWindow()
fileName = '01-02-17.xls'

# Day constants for finding sheet=day combination
MONDAY = 3
TUESDAY = 4
WEDNESDAY = 5
THURSDAY = 6
FRIDAY = 7
SATURDAY = 8
SUNDAY = 9

# Setup and copy workbook.
book = xlrd.open_workbook(filename=fileName, formatting_info=True, on_demand=True)
new_schedule = copy(book)

#Chosen day to assign. Test (will give user option to pick day) set user option to this variable.
chosen_day = TUESDAY
excelWriter = ExcelWriter(book, new_schedule.get_sheet(chosen_day), book.sheet_by_index(chosen_day), new_schedule,
                        Styles(book))

# Set Break times in new excel file.
excelWriter.calcBreakTimes()

# Set Lunch Times.
excelWriter.calcLunchTimes()
excelWriter.setYellow(6, 5)
excelWriter.setPink(5,5)

# Create editing table.
ex = VisualTable('new_schedule.xls', excelWriter.calcNumEmployees())
ex.createTable()

# Close application.
sys.exit(app.exec_())



























