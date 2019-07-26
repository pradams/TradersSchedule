import xlrd
import xlutils
import xlwt


class ExcelWriter:

    # new_book = New excel that will represent filled out schedule.
    # orig_book = The original excel file that all entries were copied from.
    def __init__(self, book, new_sheet, curr_sheet, new_book, styles):
        self.book = book
        self.new_sheet = new_sheet
        self.curr_sheet = curr_sheet
        self.new_book = new_book
        self.num_employees = self.calcNumEmployees()[0]
        self.employees = self.calcNumEmployees()[1]
        self.styles = styles
        self.open_hour = (curr_sheet.cell(1, 5).value, 5)
        self.shift_indexes = {}
        self.hour_shift_indexes = {}

    def translateHourToCell(self, time):
        hour_index = self.open_hour[1] + (round(time) - self.open_hour[0]) * 2
        return hour_index

    def calcBreakTimes(self):
        for i in range(2, self.num_employees+2):
            # Calculating actual times since cell times given in percentage of a 24 hour day.
            start_time = self.curr_sheet.cell(i, 1).value * 24
            end_time = self.curr_sheet.cell(i, 2).value * 24

            # Keep track of indexes for each shift time.
            if not self.shift_indexes.get(start_time):
                self.shift_indexes[start_time] = []
            self.shift_indexes[start_time].append(i)

            first_break = round(start_time) + 2
            second_break = round(end_time) - 2

            if first_break > 12:
                first_break -= 12
            if second_break < 0:
                second_break += 12
            elif (second_break > 12):
                second_break -= 12

            # Write calculated break times to new excel file.
            style = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin')
            self.new_sheet.write(i, 3, first_break, style)
            self.new_sheet.write(i, 4, second_break, style)

        self.new_book.save('new_schedule.xls')


    def calcLunchTimes(self):
        for time in self.shift_indexes:
            temp_count = 0
            shift_count_midpoint = round(len(self.shift_indexes[time]) / 2)
            for index in self.shift_indexes[time]:
                # Lunch time starts 4 hours after start. Create base data that connects an 8oclock lunch to index 5.
                lunch_time = round(time) + 4

                # Use base data to calculate the index for lunch time.
                lunch_index = self.open_hour[1] + (lunch_time - self.open_hour[0]) * 2
                if temp_count >= shift_count_midpoint:
                    lunch_index += 1

                # Make sure cell border styling is correct according to where in hour lunch will occur.
                if lunch_index % 2:
                    style = xlwt.easyxf('borders: left thin, top thin, bottom thin')
                else:
                    style = xlwt.easyxf('borders: right thin, top thin, bottom thin')
                self.new_sheet.write(index, int(lunch_index), 'L', style)
                temp_count += 1
        self.new_book.save('new_schedule.xls')


    # Creates list of employees. Returns number of employees and list of employees working that day.
    def calcNumEmployees(self):
        employees = []
        for i in range(2, self.curr_sheet.nrows):
            name = self.curr_sheet.cell(i, 0)
            # Check if cell type not equal to 0 (0 represents empty).
            if (name.value != ''):
                employees.append(name.value)
        return len(employees), employees


    # Set cell to yellow.
    def setYellow(self, row, col):
        style = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; borders: left thin, top thin, bottom thin;')
        self.new_sheet.write(row, col, '', style)
        style = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; borders: right thin, top thin, bottom thin;')
        self.new_sheet.write(row, col+1, '', style)
        #self.new_book.save('new_schedule.xls')

    # Set cell to pink.
    def setPink(self, row, col):
        style = xlwt.easyxf('pattern: pattern solid, fore_colour rose; borders: left thin, top thin, bottom thin;')
        self.new_sheet.write(row, col, '', style)
        style = xlwt.easyxf('pattern: pattern solid, fore_colour rose; borders: right thin, top thin, bottom thin;')
        self.new_sheet.write(row, col + 1, '', style)
        #self.new_book.save('new_schedule.xls')

    # Function returns list of employees working on specific hour. Hour should be in military time.
    def calcHourEmployees(self, hour):
        employees = []

        # open_info is a tuple with the first hour the store is open(first hour on log), and the index of the first hour.
        open_info = (int(self.open_hour[0]), 5)
        for i in range(2, self.num_employees):
            name = self.curr_sheet.cell(i, 0)

            # Extracting style information to obtain background colour of cell.
            xf = self.curr_sheet.cell_xf_index(i, (open_info[1] + (hour - open_info[0]) * 2))
            xf_next = self.book.xf_list[xf]
            colour_index = xf_next.background.pattern_colour_index

            # Checks if colour is not gray.
            if (name.value != '') and (colour_index != 22):
                employees.append(name.value)

                # Add index to dict for use in coloring the cells.
                if not self.hour_shift_indexes.get(hour):
                    self.hour_shift_indexes[hour] = []
                self.hour_shift_indexes[hour].append(i)
        return employees

    # Runs calcHourEmployees for each hour open to set up dict to be used for coloring cells.
    def setHourIndexes(self):
        for i in range(8, 21):
            self.calcHourEmployees(i)

    # Function sets up the schedule for the day, coloring the cells accordingly.
    def colorCells(self):
        self.setHourIndexes()

        # Should go up to index 28
        for i in range(8, 9):
            col = self.translateHourToCell(i)
            for row in self.hour_shift_indexes[i]:
                self.setYellow(row, col)
        self.new_book.save('new_schedule.xls')






                


