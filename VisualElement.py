from PyQt5.QtWidgets import QHeaderView, QDialog, QWidget, QPushButton, QAbstractItemView, QTableWidget, QTableWidgetItem, QVBoxLayout, QMenu, QAction
from PyQt5.QtGui import QColor, QCursor
from PyQt5.QtCore import QSize, Qt, QTimer
import datetime
import xlrd
import xlwt
from xlutils.copy import copy
import os

class VisualTable(QDialog):

    # tempSched is the newSchdule.
    def __init__(self, new_filename, numEmployees, day_index, parent=None):
        super(VisualTable, self).__init__(parent)
        self.title = "Trader's Scheduler"
        self.left = 100
        self.top = 100
        self.width = 1150
        self.height = 1000

        # Colors
        self.pink = (255, 153, 204, 255)
        self.yellow = (255, 255, 153, 255)
        self.grey = (192, 192, 192, 255)
        self.white = (0, 0, 0, 255)
        self.green = (153, 255, 153, 255)

        # Variables used for click handling (single vs double click)
        self.clock = QTimer()
        self.clock.setInterval(200)
        self.clock.setSingleShot(True)
        self.numberOfClicks = 0
        self.clock.timeout.connect(self.handleClicks)
        self.cellClicked = 0
        self.last_manager_row = 0
        self.last_manager_row_reached = False
        self.cursor_pos = 0

        # This list maintains updated rows. Lunch list maintains a pair (new_index, alignment)
        self.updatedRows = set([])
        self.updatedRowsLunch = {}
        self.original_lunch_indexes = {}

        self.read_schedule = xlrd.open_workbook(filename=new_filename, formatting_info=True, on_demand=True)
        self.write_schedule = copy(self.read_schedule)
        self.new_sheet = self.read_schedule.get_sheet(day_index+3)
        self.numberOfCE = [0] * 13
        self.numberOfPT = [0] * 13
        self.numEmployees = numEmployees
        self.save_file_name = new_filename
        self.day_index = day_index
        self.col_width = 256 * 3
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
        self.tableWidget.setRowCount(self.numEmployees[0]+3)
        self.tableWidget.setColumnCount(14)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

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
        self.tableWidget.setCellWidget(2, 0, widget)

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
            self.tableWidget.setItem(row+1, 0, cell_item)

            emp_end = end_time[3]
            emp_start = start_time[3]
            if end_time[4] != 0:
                emp_end += 0.5
            if start_time[4] != 0:
                emp_start += 0.5

            print("Test: ", (row, (emp_end-emp_start)))
            if emp_end - emp_start > 8 and not self.last_manager_row_reached:
                self.last_manager_row = row+1
            else:
                self.last_manager_row_reached = True

        print("Last: ", self.last_manager_row)
        # Set top row (title row) with hours of day.
        for col in range(1, 6):
            item = "  " + str(col+7) + "AM  "
            if str(col+7) == '12':
                item = "  " + str(col+7) + "PM  "
            self.tableWidget.setItem(2, col, QTableWidgetItem(item))
            head.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        for col in range(6, 14):
            item = "  " + str(col - 5) + "PM  "
            self.tableWidget.setItem(2, col, QTableWidgetItem(item))
            head.setSectionResizeMode(col, QHeaderView.ResizeToContents)

        # Resize cell to length of name.
        head.setSectionResizeMode(0, QHeaderView.ResizeToContents)

        # Copy all the cell colors from the edited excel file. Also, connect save button to action.
        self.copyCellColors()
        self.tableWidget.clicked.connect(self.clickedCell)
        #self.tableWidget.doubleClicked.connect(self.handleDoubleClick)

        # Set CE/PT Count row a different color than the rest, and connect to numbers stored (self.numOfCE/self.numOfPT)
        for col in range(0, self.tableWidget.columnCount()):
            if col == 0:
                cell_item_CE = QTableWidgetItem('# of CE Members: ')
                cell_item_PT = QTableWidgetItem('# of PT Members: ')
                cell_item_CE.setTextAlignment(Qt.AlignCenter)
                cell_item_PT.setTextAlignment(Qt.AlignCenter)
            else:
                item_CE = "   " + str(self.numberOfCE[col-1]) + " "
                cell_item_CE = QTableWidgetItem(item_CE)
                item_PT = "   " + str(self.numberOfPT[col-1]) + " "
                cell_item_PT = QTableWidgetItem(item_PT)
            cell_item_CE.setBackground(QColor(255, 255, 153))
            cell_item_PT.setBackground(QColor(255, 153, 204))

            self.tableWidget.setItem(0, col, cell_item_CE)
            self.tableWidget.setItem(1, col, cell_item_PT)

        # Set tablewidget to QDialog.
        self.setLayout(self.layout)
        self.exec_()
        self.tableWidget.show()
        pass

    # Copies all cell colors from the newly colored excel file into the editing table.
    def copyCellColors(self):
        # Colour indexes ----- Pink: 45    Yellow: 43
        col = 1

        # Iterate over every cell, col first, row second.
        for i in range(5, 30, 2):
            for row in range(2, self.numEmployees[0] + 2):
                # Extracting style information to obtain background colour of cell.
                xf_left = self.new_sheet.cell_xf_index(row, i)
                xf_next_left = self.read_schedule.xf_list[xf_left]
                colour_index_left = xf_next_left.background.pattern_colour_index

                xf_right = self.new_sheet.cell_xf_index(row, i+1)
                xf_next_right = self.read_schedule.xf_list[xf_right]
                colour_index_right = xf_next_right.background.pattern_colour_index

                if (self.new_sheet.cell(row, i).value == 'L'):
                    cell_item = QTableWidgetItem('L')
                    cell_item.setTextAlignment(1)
                    self.original_lunch_indexes[row+1] = col
                elif (self.new_sheet.cell(row, i+1).value == 'L'):
                    cell_item = QTableWidgetItem('L')
                    cell_item.setTextAlignment(2)
                    self.original_lunch_indexes[row + 1] = col
                else:
                    cell_item = QTableWidgetItem('')

                cell_item.setSizeHint(QSize(2, 2))

                # Set the table cell the same color as the excel cell.
                if colour_index_left == colour_index_right:
                    if colour_index_left == 45:
                        cell_item.setBackground(QColor(self.pink[0], self.pink[1], self.pink[2]))
                        if cell_item.text() != 'L':
                            self.numberOfPT[col-1] += 1
                    elif colour_index_left == 43:
                        cell_item.setBackground(QColor(self.yellow[0], self.yellow[1], self.yellow[2]))
                        self.numberOfCE[col-1] += 1
                    elif colour_index_left == 22:
                        cell_item.setBackground(QColor(self.grey[0], self.grey[1], self.grey[2]))
                    self.tableWidget.setItem(row+1, col, cell_item)
                else:
                    cell_item.setBackground(QColor(self.grey[0], self.grey[1], self.grey[2]))
                    self.tableWidget.setItem(row+1, col, cell_item)
                '''

                else:
                    cell_item.setBackground(QColor(self.grey[0], self.grey[1], self.grey[2]))
                    self.tableWidget.setItem(row, col, cell_item)
                '''

            col += 1

    # Saves any altered cells back to the excel file. Called when save button is clicked.
    def saveToExcel(self, write_sheet):

        # Update excel file with any changed lunches.
        for new_lunch_row in self.updatedRowsLunch:
            self.setLunch((new_lunch_row, self.updatedRowsLunch[new_lunch_row][0]), self.updatedRowsLunch[new_lunch_row][1],
                          write_sheet)

        # Update the excel file with all updated cells in editor.
        for index in self.updatedRows:
            cell_color = self.tableWidget.item(index[0], index[1]).background().color().getRgb()
            print("Cell color: ", cell_color)
            print("Index: ", index)
            if cell_color == self.pink:
                print("Setting Pink")
                self.setPink(index, write_sheet)
            elif cell_color == self.green:
                self.setGreen(index, write_sheet)
            else:
                self.setYellow(index, write_sheet)

        # Print CE count numbers in excel file.
        style_main_label = xlwt.easyxf('pattern: pattern solid, fore_colour white; borders: left thin, right thin, top thin, bottom thin')
        write_sheet.write(self.numEmployees[0]+3, 0, 'Number of CE Members', style_main_label)
        style_left = xlwt.easyxf('pattern: pattern solid, fore_colour white; borders: left thin, top thin, bottom thin; align: horiz right')
        style_right = xlwt.easyxf('pattern: pattern solid, fore_colour white; borders: right thin, top thin, bottom thin; align: horiz right')
        ce_list_index = 0
        
        for col in range(5, 30, 2):
            write_sheet.col(col).width = self.col_width
            write_sheet.write(self.numEmployees[0]+3, col, self.numberOfCE[ce_list_index], style_left)
            write_sheet.write(self.numEmployees[0]+3, col+1, '', style_right)
            ce_list_index += 1

        print("Saving to File")
    # Method helps distinguish between single and double clicks
    def clickedCell(self, cell):
        self.numberOfClicks += 1
        self.cellClicked = cell
        if not self.clock.isActive():
            self.clock.start()

    # Handle the cell being clicked.
    def handleClicks(self):
        try:
            row = self.cellClicked.row()
            col = self.cellClicked.column()
            clicked_cell = self.tableWidget.item(row, col)
            current_color = clicked_cell.background().color().getRgb()

            if self.tableWidget.item(row, col).text() == '':
                if self.numberOfClicks == 1:
                    if current_color == self.grey:
                        pass

                    elif current_color == self.pink:
                        clicked_cell.setBackground(QColor(self.yellow[0], self.yellow[1],
                                                              self.yellow[2]))
                        self.updatedRows.add((row, col))

                        # Checks if the cell belonged to manager to see if it should be included in CE Count.
                        if row > self.last_manager_row:
                            self.numberOfCE[col - 1] += 1
                            self.numberOfPT[col - 1] -= 1
                    elif current_color == self.yellow:
                        clicked_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                                              self.pink[2]))
                        self.updatedRows.add((row, col))
                        if row > self.last_manager_row:
                            self.numberOfCE[col - 1] -= 1
                            self.numberOfPT[col - 1] += 1
                    else:
                        clicked_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                                              self.pink[2]))
                        self.updatedRows.add((row, col))
                        if row > self.last_manager_row:
                            self.numberOfPT[col - 1] += 1

                elif self.numberOfClicks == 2:
                    if current_color == self.grey:
                        pass

                    else:
                        clicked_cell.setBackground(QColor(self.green[0], self.green[1],
                                                              self.green[2]))
                        if current_color == self.yellow:
                            if row > self.last_manager_row:
                                self.numberOfCE[col - 1] -= 1
                        elif current_color == self.pink:
                            if row > self.last_manager_row:
                                self.numberOfPT[col-1] -= 1

                        self.updatedRows.add((row, col))

                self.updateCountLabels(col)
                self.numberOfClicks = 0

        except Exception as e:
            print(e)
            pass

    # Function updates the cell count depending on color change.
    def updateCountLabels(self, col):
        item_CE = "   " + str(self.numberOfCE[col - 1]) + "   "
        cell_item_CE = QTableWidgetItem(item_CE)
        cell_item_CE.setBackground(QColor(255, 255, 153))
        self.tableWidget.setItem(0, col, cell_item_CE)

        item_PT = "   " + str(self.numberOfPT[col - 1]) + "   "
        cell_item_PT = QTableWidgetItem(item_PT)
        cell_item_PT.setBackground(QColor(255, 153, 204))
        self.tableWidget.setItem(1, col, cell_item_PT)

    # Set cell to yellow.
    def setYellow(self, index, write_sheet):
        left_value = self.new_sheet.cell(index[0]-1, (index[1] * 2) + 3).value
        right_value = self.new_sheet.cell(index[0]-1, (index[1] * 2) + 4).value
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_yellow; borders: left thin, top thin, bottom thin;')
        write_sheet.write(index[0]-1, (index[1] * 2) + 3, left_value, style)
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_yellow; borders: right thin, top thin, bottom thin;')
        write_sheet.write(index[0]-1, (index[1] * 2) + 4, right_value, style)
        # self.new_book.save('new_schedule.xls')

    # Set cell to pink.
    def setPink(self, index, write_sheet):
        left_value = self.new_sheet.cell(index[0]-1, (index[1] * 2) + 3).value
        right_value = self.new_sheet.cell(index[0]-1, (index[1] * 2) + 4).value
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour rose; borders: left thin, top thin, bottom thin;')
        write_sheet.write(index[0]-1, (index[1] * 2) + 3, left_value, style)
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour rose; borders: right thin, top thin, bottom thin;')
        write_sheet.write(index[0]-1, (index[1] * 2) + 4, right_value, style)
        # self.new_book.save('new_schedule.xls')

    # Set cell color to green for bridge.
    def setGreen(self, index, write_sheet):
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_green; borders: left thin, top thin, bottom thin;')
        write_sheet.write(index[0]-1, (index[1] * 2) + 3, '', style)
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_green; borders: right thin, top thin, bottom thin;')
        write_sheet.write(index[0]-1, (index[1] * 2) + 4, '', style)

    # Sets lunch in excel file.
    def setLunch(self, index, top, write_sheet):
        style_left = xlwt.easyxf(
            'pattern: pattern solid, fore_colour rose; borders: left thin, top thin, bottom thin;')
        style_right = xlwt.easyxf(
            'pattern: pattern solid, fore_colour rose; borders: right thin, top thin, bottom thin;')

        # Get rid of original lunch.
        write_sheet.write(index[0] - 1, (self.original_lunch_indexes[index[0]] * 2) + 3, '', style_left)
        write_sheet.write(index[0] - 1, (self.original_lunch_indexes[index[0]] * 2) + 4, '', style_right)

        if top:
            write_sheet.write(index[0] - 1, (index[1] * 2) + 3, 'L', style_left)
            write_sheet.write(index[0] - 1, (index[1] * 2) + 4, '', style_right)
        else:
            write_sheet.write(index[0] - 1, (index[1] * 2) + 3, '', style_left)
            write_sheet.write(index[0] - 1, (index[1] * 2) + 4, 'L', style_right)

    # Handle situation where save button is clicked.
    # Should update the new schedule with edited cell colors.
    def save_button_clicked(self):
        write_sheet = self.write_schedule.get_sheet(self.day_index+3)
        self.saveToExcel(write_sheet)
        self.read_schedule.release_resources()
        os.remove(self.save_file_name)
        self.write_schedule.save(self.save_file_name)
        print("Just Wrote")
        self.close()

    # Sets up right click menu.
    def contextMenuEvent(self, event):
        self.menu = QMenu(self)

        leftLunchAction = QAction('Change Lunch To Top of Hour', self)
        leftLunchAction.triggered.connect(lambda: self.changeToTopHourLunch(event))

        rightLunchAction = QAction('Change Lunch To Bottom of Hour', self)
        rightLunchAction.triggered.connect(lambda: self.changeToBottomHourLunch(event))

        print("New: ", event)
        self.cursor_pos = QCursor.pos()
        self.menu.addAction(leftLunchAction)
        self.menu.addAction(rightLunchAction)
        self.menu.popup(self.cursor_pos)

    # Both methods handle relative option chosen from right click menu.
    def changeToTopHourLunch(self, event):
        row = self.tableWidget.rowAt(self.cursor_pos.y())-4
        col = self.tableWidget.columnAt(self.cursor_pos.x())-2
        self.deletePreviousLunch(row)

        # Now put lunch at top of hour of selected cell. Add to list of updated lunches.
        self.tableWidget.item(row, col).setText('L')
        self.tableWidget.item(row, col).setTextAlignment(1)

        self.colorPinkForLunch(row, col)

        # Add new_index/ alignment pair to updatedRowsLunch dict. True stands for left(top), False stands for right(bottom)
        self.updatedRowsLunch[row] = (col, True)

    def changeToBottomHourLunch(self, event):
        row = self.tableWidget.rowAt(self.cursor_pos.y())-4
        col = self.tableWidget.columnAt(self.cursor_pos.x())-2
        self.deletePreviousLunch(row)

        # Now put lunch at top of hour of selected cell. Add to list of updated lunches.
        self.tableWidget.item(row, col).setText('L')
        self.tableWidget.item(row, col).setTextAlignment(2)

        self.colorPinkForLunch(row, col)

        # Add pair to updatedRowsLunch. See function above for details.
        self.updatedRowsLunch[row] = (col, False)

    def colorPinkForLunch(self, row, col):
        current_cell = self.tableWidget.item(row,col)
        cell_color = current_cell.background().color().getRgb()
        if cell_color == self.yellow:
            print("Changing to Pink")
            current_cell.setBackground(QColor(self.pink[0], self.pink[1],
                                              self.pink[2]))
            self.numberOfCE[col-1] -= 1
            self.numberOfPT[col-1] += 1
            self.updateCountLabels(col)



    def deletePreviousLunch(self, row):
        # Iterate through columns to find last lunch position. Clear this position.
        for col in range(1, self.tableWidget.columnCount()):
            if self.tableWidget.item(row, col).text() == 'L':
                self.tableWidget.item(row, col).setText('')


