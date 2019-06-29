from PyQt5.QtWidgets import QMainWindow, QApplication, QHeaderView, QWidget, QAction, QTableWidget, QTableWidgetItem, QVBoxLayout
from PyQt5.QtGui import QIcon, QBrush, QColor
from PyQt5.QtCore import pyqtSlot, QSize

class VisualElement(QWidget):

    def __init__(self, tempSched, numEmployees):
        super().__init__()
        self.title = "Trader's Scheduler"
        self.left = 100
        self.top = 100
        self.width = 1000
        self.height = 1000
        self.tempSched = tempSched
        self.numEmployees = numEmployees
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.createTable()

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)
        self.show()

    def createTable(self):
        self.tableWidget = QTableWidget()
        self.tableWidget.setRowCount(self.numEmployees[0]+1)
        self.tableWidget.setColumnCount(14)

        sheet = self.tempSched.sheet_by_index(4)

        # Set first column as employee names.
        for row in range(1, self.numEmployees[0]+1):
            self.tableWidget.setItem(row, 0, QTableWidgetItem(self.numEmployees[1][row-1]))

        head = self.tableWidget.horizontalHeader()
        head.setSectionResizeMode(0, QHeaderView.ResizeToContents)

        for row in range(1, self.numEmployees[0]+1):
            cell_item = QTableWidgetItem('')
            cell_item.setBackground(QColor(243, 159, 255))
            cell_item.setSizeHint(QSize(2,2))
            #self.tableWidget.setItem(row, 1, QTableWidgetItem().setForeground(QColor(235, 100, 255)))
            self.tableWidget.setItem(row, 1, cell_item)

        self.tableWidget.move(0,0)


