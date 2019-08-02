import xlrd
import sys
from VisualElement import VisualTable
from MainWindow import MainIntroWindow
from ExcelWriter import ExcelWriter
from xlutils.copy import copy
from xlutils.styles import Styles
from PyQt5.QtWidgets import QApplication


### Create Applicatin Window ###
app = QApplication(sys.argv)

# Create filebrowser and day selector.
browser = MainIntroWindow()

# Close application.
app.exec()
#sys.exit(app.exec_())



























