from PyQt5.QtWidgets import QMainWindow, QApplication, QHeaderView, QLabel, QWidget, QPushButton, QTableWidget, QTableWidgetItem, QGridLayout, QVBoxLayout, QComboBox, QFileDialog
from PyQt5.QtGui import QIcon, QBrush, QColor, QPalette
from PyQt5.QtCore import pyqtSlot, QSize, QStringListModel
import xlrd
import xlwt
from xlutils.copy import copy

class VisualTable(QWidget):

    # tempSched is the newSchdule.
    def __init__(self, new_filename, numEmployees):
        super().__init__()
        self.title = "Trader's Scheduler"
        self.left = 100
        self.top = 100
        self.width = 750
        self.height = 700

        # Colors
        self.pink = (255, 153, 204, 255)
        self.yellow = (255, 255, 153, 255)
        self.grey = (192, 192, 192, 255)

        self.read_schedule = xlrd.open_workbook(filename=new_filename, formatting_info=True, on_demand=True)
        self.write_schedule = copy(self.read_schedule)
        self.numEmployees = numEmployees
        self.updatedRows = set([])
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.tableWidget = QTableWidget()
        self.tableWidget.setMouseTracking(True)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)
        self.current_cell = [0,0]
        self.show()

    def createTable(self):
        self.tableWidget.setRowCount(self.numEmployees[0]+1)
        self.tableWidget.setColumnCount(14)

        # Hide the vertical and horizontal indexes.
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)
        self.tableWidget.setSelectionMode(QTableWidget.NoSelection)

        # Create button
        save_button = QPushButton(self.tableWidget)
        save_button.setText("Save Schedule")
        palette = save_button.palette()
        palette.setColor(QPalette.Button, QColor('blue'))
        save_button.setPalette(palette)
        save_button.update()
        save_button.clicked.connect(self.save_button_clicked)

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

        self.copyCellColors(self.read_schedule.get_sheet(4))
        '''
        for row in range(1, self.numEmployees[0]+1):
            cell_item = QTableWidgetItem('')
            cell_item.setBackground(QColor(self.pink[0], self.pink[1], self.pink[2]))
            cell_item.setSizeHint(QSize(2,2))
            self.tableWidget.setItem(row, 1, cell_item)
        '''
        self.tableWidget.clicked.connect(self.clickedCell)


    def copyCellColors(self, new_sheet):

        # Colour indexes ----- Pink: 45    Yellow: 43
        col = 1
        for i in range(5, 30, 2):
            for row in range(2, self.numEmployees[0] + 2):
                # Extracting style information to obtain background colour of cell.
                xf = new_sheet.cell_xf_index(row, i)
                xf_next = self.read_schedule.xf_list[xf]
                colour_index = xf_next.background.pattern_colour_index

                cell_item = QTableWidgetItem('')
                cell_item.setSizeHint(QSize(2, 2))

                if colour_index == 45:
                    cell_item.setBackground(QColor(self.pink[0], self.pink[1], self.pink[2]))
                    print("Setting Pink")
                elif colour_index == 43:
                    cell_item.setBackground(QColor(self.yellow[0], self.yellow[1], self.yellow[2]))
                    print("Setting Yellow")
                elif colour_index == 22:
                    cell_item.setBackground(QColor(self.grey[0], self.grey[1], self.grey[2]))
                    print("Setting Grey")

                self.tableWidget.setItem(row-1, col, cell_item)
            col += 1


    # Handle the cell being clicked.
    def clickedCell(self, cell):
        try:
            row = cell.row()
            col = cell.column()
            clicked_cell = self.tableWidget.item(row, col)
            current_color = clicked_cell.background().color().getRgb()
            if current_color == self.pink:
                clicked_cell.setBackground(QColor(self.yellow[0], self.yellow[1],
                                                     self.yellow[2]))
                self.updatedRows.add(row)
            elif current_color == self.yellow:
                clicked_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                              self.pink[2]))
                self.updatedRows.add(row)
        except:
            print("Blank cell chosen. No Color Change.")
            pass


    # Handle situation where save button is clicked.
    # Should update the new schedule with edited cell colors.
    def save_button_clicked(self):
        print("Button Pressed")


class MainIntroWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Trader's Scheduler: File Browser")
        self.setGeometry(30, 30, 300, 300)

        self.layout = QGridLayout(self)

        # Initialize all widgets
        self.day_selector = QComboBox()
        self.day_selector.addItems(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
                                    "Friday", "Saturday", "All"])
        self.choose_day_label = QLabel("Choose Day: ")
        self.open_file_button = QPushButton("Open Previous Schedule")
        self.open_file_button.clicked.connect(self.openFileBrowser)

        self.layout.addWidget(self.choose_day_label, 0, 0)
        self.layout.addWidget(self.day_selector, 0, 1)
        self.layout.addWidget(self.open_file_button, 1, 0)

        self.setLayout(self.layout)
        self.show()

    def openFileBrowser(self):
        file_dialog = QFileDialog()
        file_dialog.move(200, 200)
        options = file_dialog.Options()
        options |= file_dialog.DontUseNativeDialog

        fileName, _ = file_dialog.getOpenFileName(self, "Choose a schedule", "", "All Files (*)", options=options)
        return fileName







