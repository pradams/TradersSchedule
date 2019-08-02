from PyQt5.QtWidgets import QHeaderView, QLabel, QDialog, QWidget, QPushButton, QMainWindow, QTableWidget, QTableWidgetItem, QGridLayout, QVBoxLayout, QComboBox, QFileDialog
from PyQt5.QtGui import QColor, QPalette, QFont
from PyQt5.QtCore import QSize, Qt
import datetime
from ExcelWriter import ExcelWriter
from xlutils.copy import copy
from xlutils.styles import Styles


import xlrd
import xlwt
from xlutils.copy import copy

class VisualTable(QDialog):

    # tempSched is the newSchdule.
    def __init__(self, new_filename, numEmployees, day_index, parent=None):
        super(VisualTable, self).__init__(parent)
        self.title = "Trader's Scheduler"
        self.left = 100
        self.top = 100
        self.width = 900
        self.height = 700

        # Colors
        self.pink = (255, 153, 204, 255)
        self.yellow = (255, 255, 153, 255)
        self.grey = (192, 192, 192, 255)
        self.white = (0, 0, 0, 255)

        self.read_schedule = xlrd.open_workbook(filename=new_filename, formatting_info=True, on_demand=True)
        self.write_schedule = copy(self.read_schedule)
        self.new_sheet = self.read_schedule.get_sheet(day_index+3)
        self.numberOfCE = [0] * 13
        self.numEmployees = numEmployees
        self.updatedRows = set([])
        self.save_file_name = new_filename
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.tableWidget = QTableWidget(self)
        self.tableWidget.setMouseTracking(True)
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)
        self.current_cell = [0,0]
        self.tableWidget.show()
        #self.show()

    def createTable(self):
        print("Creating Table")
        # Settings for main table with employees
        self.tableWidget.setRowCount(self.numEmployees[0]+2)
        self.tableWidget.setColumnCount(14)

        # Hide the vertical and horizontal indexes.
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)
        self.tableWidget.setSelectionMode(QTableWidget.NoSelection)
        #self.tableWidget.setStyleSheet("border-radius: 2px;")

        # Create button
        widget = QWidget()
        layout = QVBoxLayout(widget)
        save_button = QPushButton(self.tableWidget)
        save_button.setText("    Save Schedule   ")

        save_button.setStyleSheet("border-radius: 0.5px;"
                                  "background-color: gray;"
                                  "color: white;"
                                  "padding: 3px 25px 3px 25px;")

        save_button.update()
        save_button.clicked.connect(self.save_button_clicked)
        layout.addWidget(save_button)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        widget.setLayout(layout)
        self.tableWidget.setCellWidget(1, 0, widget)

        head = self.tableWidget.horizontalHeader()

        # Set first column as employee names and start and end times.
        for row in range(2, self.numEmployees[0]+2):

            # Extract the time from the excel document and format correctly.
            start_time = xlrd.xldate_as_tuple(self.new_sheet.cell(row,1).value, self.read_schedule.datemode)
            end_time = xlrd.xldate_as_tuple(self.new_sheet.cell(row,2).value, self.read_schedule.datemode)
            date_string_start = str(start_time[3]) + ":" + str(start_time[4])
            date_string_end = str(end_time[3]) + ":" + str(end_time[4])
            end_date = datetime.datetime.strptime(date_string_end, '%H:%M').strftime('%I:%M %p')
            start_date = datetime.datetime.strptime(date_string_start, '%H:%M').strftime('%I:%M %p')
            item = self.numEmployees[1][row-2] + "  |  " + start_date + "-" + end_date
            cell_item = QTableWidgetItem(item)
            if row % 2 != 0:
                cell_item.setBackground(QColor(235,235,235))
            self.tableWidget.setItem(row, 0, cell_item)

        # Set top row (title row) with hours of day.
        for col in range(1, 6):
            item = "  " + str(col+7) + "  "
            self.tableWidget.setItem(1, col, QTableWidgetItem(item))
            head.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        for col in range(6, 14):
            item = "  " + str(col - 5) + "  "
            self.tableWidget.setItem(1, col, QTableWidgetItem(item))
            head.setSectionResizeMode(col, QHeaderView.ResizeToContents)

        # Resize cell to length of name.
        head.setSectionResizeMode(0, QHeaderView.ResizeToContents)

        # Copy all the cell colors from the edited excel file. Also, connect save button to action.
        self.copyCellColors()
        self.tableWidget.clicked.connect(self.clickedCell)

        # Set CE Count row a different color than the rest, and connect to numbers stored (self.numOfCE)
        for col in range(0, self.tableWidget.columnCount()):
            if col == 0:
                cell_item = QTableWidgetItem('# of CE Members: ')
                cell_item.setTextAlignment(Qt.AlignCenter)
            else:
                item = "  " + str(self.numberOfCE[col-1]) + "  "
                cell_item = QTableWidgetItem(item)
            cell_item.setBackground(QColor(115, 194, 246))
            self.tableWidget.setItem(0, col, cell_item)


        # Set tablewidget to QDialog.
        self.setLayout(self.layout)
        self.exec_()
        self.tableWidget.show()


    def copyCellColors(self):
        # Colour indexes ----- Pink: 45    Yellow: 43
        col = 1

        # Iterate over every cell, col first, row second.
        for i in range(5, 30, 2):
            for row in range(2, self.numEmployees[0] + 2):
                # Extracting style information to obtain background colour of cell.
                xf = self.new_sheet.cell_xf_index(row, i)
                xf_next = self.read_schedule.xf_list[xf]
                colour_index = xf_next.background.pattern_colour_index

                if (self.new_sheet.cell(row, i).value == 'L'):
                    cell_item = QTableWidgetItem('L')
                    cell_item.setTextAlignment(1)
                elif (self.new_sheet.cell(row, i+1).value == 'L'):
                    cell_item = QTableWidgetItem('L')
                    cell_item.setTextAlignment(2)
                else:
                    cell_item = QTableWidgetItem('')


                cell_item.setSizeHint(QSize(2, 2))

                # Set the table cell the same color as the excel cell.
                if colour_index == 45:
                    cell_item.setBackground(QColor(self.pink[0], self.pink[1], self.pink[2]))
                elif colour_index == 43:
                    cell_item.setBackground(QColor(self.yellow[0], self.yellow[1], self.yellow[2]))
                    self.numberOfCE[col-1] += 1
                elif colour_index == 22:
                    cell_item.setBackground(QColor(self.grey[0], self.grey[1], self.grey[2]))

                self.tableWidget.setItem(row, col, cell_item)
            col += 1


    def saveToExcel(self, write_sheet):
        for index in self.updatedRows:
            cell_color = self.tableWidget.item(index[0], index[1]).background().color().getRgb()

            if cell_color == self.pink:
                self.setPink(index, write_sheet)
            else:
                self.setYellow(index, write_sheet)

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
                self.updatedRows.add((row, col))
                self.numberOfCE[col-1] += 1
            elif current_color == self.yellow:
                clicked_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                              self.pink[2]))
                self.updatedRows.add((row, col))
                self.numberOfCE[col-1] -= 1
            elif current_color == self.white:
                clicked_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                                  self.pink[2]))
                self.updatedRows.add((row, col))

            self.updateCECountLabels(col)
        except Exception as e:
            print(e)
            pass

    # Function updates the cell count depending whether changed to yellow or from yellow.
    def updateCECountLabels(self, col):
        item = "  " + str(self.numberOfCE[col - 1]) + "  "
        cell_item = QTableWidgetItem(item)
        #cell_item.setFont(QFont("Helvetica", 10, QFont().bold))
        cell_item.setBackground(QColor(115, 194, 246))
        self.tableWidget.setItem(0, col, cell_item)

    # Set cell to yellow.
    def setYellow(self, index, write_sheet):
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_yellow; borders: left thin, top thin, bottom thin;')
        write_sheet.write(index[0] + 1, (index[1] * 2) + 3, '', style)
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_yellow; borders: right thin, top thin, bottom thin;')
        write_sheet.write(index[0] + 1, (index[1] * 2) + 4, '', style)
        # self.new_book.save('new_schedule.xls')

    # Set cell to pink.
    def setPink(self, index, write_sheet):
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour rose; borders: left thin, top thin, bottom thin;')
        write_sheet.write(index[0] + 1, (index[1] * 2) + 3, '', style)
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour rose; borders: right thin, top thin, bottom thin;')
        write_sheet.write(index[0] + 1, (index[1] * 2) + 4, '', style)
        # self.new_book.save('new_schedule.xls')


    # Handle situation where save button is clicked.
    # Should update the new schedule with edited cell colors.
    def save_button_clicked(self):
        write_sheet = self.write_schedule.get_sheet(4)
        self.saveToExcel(write_sheet)
        self.write_schedule.save(self.save_file_name)
        print("Button Pressed")
        self.close()











