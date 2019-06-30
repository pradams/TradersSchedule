from PyQt5.QtWidgets import QMainWindow, QApplication, QHeaderView, QWidget, QAction, QTableWidget, QTableWidgetItem, QVBoxLayout
from PyQt5.QtGui import QIcon, QBrush, QColor
from PyQt5.QtCore import pyqtSlot, QSize
from pynput import keyboard

class VisualElement(QWidget):

    def __init__(self, tempSched, numEmployees):
        super().__init__()
        self.title = "Trader's Scheduler"
        self.left = 100
        self.top = 100
        self.width = 750
        self.height = 700
        self.tempSched = tempSched
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
            cell_item.setBackground(QColor(243, 159, 255))
            cell_item.setSizeHint(QSize(2,2))
            self.tableWidget.setItem(row, 1, cell_item)

        self.tableWidget.cellEntered.connect(self.enteredCell)

        self.tableWidget.move(0,0)

    def enteredCell(self, row, col):
        data = self.tableWidget.item(row, col)

        pass


    '''
    def on_press(self, key):
        print('{0} pressed'.format(key))

    def on_release(self, key):
        if key == Key.esc:
            return False

    listener = keyboard.Listener(
        on_press=on_press,
        on_release=on_release)
    listener.start()

    ''' 

