from PyQt5.QtWidgets import QMainWindow, QApplication, QHeaderView, QWidget, QAction, QTableWidget, QTableWidgetItem, QVBoxLayout
from PyQt5.QtGui import QIcon, QBrush, QColor
from PyQt5.QtCore import pyqtSlot, QSize
import xlrd
import xlwt
from xlutils.copy import copy

class VisualElement(QWidget):

    # tempSched is the newSchdule.
    def __init__(self, new_filename, numEmployees):
        super().__init__()
        self.title = "Trader's Scheduler"
        self.left = 100
        self.top = 100
        self.width = 750
        self.height = 700

        # Colors
        self.pink = (243, 159, 255, 255)
        self.yellow = (255, 255, 97, 255)

        self.read_schedule = xlrd.open_workbook(filename=new_filename, formatting_info=True, on_demand=True)
        self.write_schedule = copy(self.read_schedule)
        self.numEmployees = numEmployees
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.tableWidget = QTableWidget()
        self.tableWidget.setMouseTracking(True)
        self.createTable()

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)
        self.current_cell = [0,0]
        self.show()

    def createTable(self):
        self.tableWidget.setRowCount(self.numEmployees[0]+1)
        self.tableWidget.setColumnCount(14)

        # Hide the vertical and horizontal indexes.
        self.tableWidget.verticalHeader().setVisible = False
        self.tableWidget.horizontalHeader().setVisible = False
        self.tableWidget.setSelectionMode(QTableWidget.NoSelection)

        head = self.tableWidget.horizontalHeader()

        # Set first column as employee names.
        for row in range(1, self.numEmployees[0]+1):
            self.tableWidget.setItem(row, 0, QTableWidgetItem(self.numEmployees[1][row-1]))

        # Set top row (title row) with hours of day.
        for col in range(1, 6):
            item = "  " + str(col+7) + "  "
            self.tableWidget.setItem(0, col, QTableWidgetItem(item))
            head.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        for col in range(6, 14):
            item = "  " + str(col - 5) + "  "
            self.tableWidget.setItem(0, col, QTableWidgetItem(item))
            head.setSectionResizeMode(col, QHeaderView.ResizeToContents)

        # Resize cell to length of name.
        head.setSectionResizeMode(0, QHeaderView.ResizeToContents)

        for row in range(1, self.numEmployees[0]+1):
            cell_item = QTableWidgetItem('')
            cell_item.setBackground(QColor(self.pink[0], self.pink[1], self.pink[2]))
            cell_item.setSizeHint(QSize(2,2))
            self.tableWidget.setItem(row, 1, cell_item)

        self.tableWidget.clicked.connect(self.clickedCell)


    # Handle the cell being clicked.
    def clickedCell(self, cell):
        row = cell.row()
        col = cell.column()
        clicked_cell = self.tableWidget.item(row, col)
        current_color = clicked_cell.background().color().getRgb()
        if current_color == self.pink:
            clicked_cell.setBackground(QColor(self.yellow[0], self.yellow[1],
                                                     self.yellow[2]))
        elif current_color == self.yellow:
            clicked_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                              self.pink[2]))


