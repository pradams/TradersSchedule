import xlrd
import xlutils
import xlwt
import math
import random


class ExcelWriter:

    # new_book = New excel that will represent filled out schedule.
    # orig_book = The original excel file that all entries were copied from.
    def __init__(self, book, new_sheet, curr_sheet, new_book, styles, save_file_name):
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
        self.save_file_name = save_file_name
        self.col_width = 256 * 3
        self.last_manager_row = 0

        # Contains the indexes that should have an L (for lunch). Used for writing L's after coloring cells.
        self.lunch_indexes = []

        #Lists for sections of shifts.
        self.open_shifts = []
        self.morning_shifts = []
        self.mid_shifts = []
        self.closing_shifts = []
        self.all_shifts = []

        #Recommended amount of CE members for each hour.
        self.recommended_ce = [6, 9, 10, 11, 11, 11, 10, 11, 12, 13, 13, 11, 9]

        # Reference matrix to store info for all cells.
        self.reference_matrix = []

        ####### USED FOR TESTING ################################################################################
        # Array holds different hour assignments for each shift.
        # [0: 5:00, 1: 5:30, 2: 6:00, 3: 6:00, 4: 6:30, 5: 7:00, 6: 7:30, 7: 10:00, 8: 11:00, 9: 12:00, 10: 12:00
        # 11: 1:00, 12: 1:30, 13: 2:00, 14: 2:30, 15: 2:30, 16: 9:00, 17: 8:00]

        self.shift_assignments = [['pink', 'yellow', 'yellow', 'pink', 'pink'],
                             ['pink', 'yellow', 'yellow', 'pink', 'pink', 'pink'],
                             ['pink', 'pink', 'pink', 'yellow', 'yellow', 'pink'],
                             ['pink', 'yellow', 'pink', 'yellow', 'pink', 'pink'],
                             ['yellow', 'pink', 'pink', 'yellow', 'pink', 'yellow', 'pink'],
                             ['yellow', 'pink', 'yellow', 'pink', 'pink', 'pink', 'yellow'],
                             ['yellow', 'yellow', 'yellow', 'yellow', 'yellow', 'yellow', 'yellow'],
                             ['yellow', 'pink', 'pink', 'pink', 'pink', 'yellow', 'pink', 'yellow'],
                             ['yellow', 'pink', 'pink', 'yellow', 'pink', 'pink', 'pink', 'yellow'],
                             ['pink', 'yellow', 'yellow', 'pink', 'pink', 'pink', 'yellow', 'pink'],
                             ['yellow', 'yellow', 'pink', 'pink', 'pink', 'yellow', 'pink', 'yellow'],
                             ['yellow', 'pink', 'pink', 'pink', 'pink', 'yellow', 'pink', 'yellow'],
                             ['pink', 'yellow', 'yellow', 'pink', 'pink', 'pink', 'pink', 'yellow'],
                             ['yellow', 'yellow', 'pink', 'pink', 'pink', 'pink', 'yellow', ''],
                             ['pink', 'pink', 'yellow', 'pink', 'yellow', 'pink', 'pink', ''],
                             ['pink', 'yellow', 'pink', 'pink', 'yellow', 'yellow', 'pink', ''],
                             ['yellow', 'pink', 'yellow', 'pink', 'pink', 'yellow', 'pink', 'yellow'],
                             ['yellow', 'pink', 'yellow', 'pink', 'pink', 'yellow', 'pink', 'yellow']]

        self.shift_assignment_indices = {5.0: 0, 5.5: 1, 6.0: 2, 6.5: 4, 7.0: 5, 7.5: 6, 8.0: 17, 9.0: 16, 10.0: 7, 11.0: 8, 12.0: 9, 13.0: 11, 13.5: 12,
                         14.0: 13, 14.5: 14}

        ##########################################################################################################

    def translateHourToCell(self, time):
        if time < 8:
            hour_index = self.open_hour[1] + (math.ceil(time) - self.open_hour[0]) * 2
        else:
            hour_index = self.open_hour[1] + (math.ceil(time) - self.open_hour[0]) * 2
        return hour_index

    def calcBreakTimes(self):
        for row in range(2, self.num_employees+2):
            # Calculating actual times since cell times given in percentage of a 24 hour day.
            start_time = self.curr_sheet.cell(row, 1).value * 24
            end_time = self.curr_sheet.cell(row, 2).value * 24

            # If manager (greater than 8 hours), record row. If row greater than certain number, don't change
            # because it might be merchant that was not moved to top of log.
            if end_time - start_time > 8:
                if row < 18:
                    self.last_manager_row = row

            # Keep track of indexes for each shift time.
            if end_time - start_time > 5:
                if not self.shift_indexes.get(start_time):
                    self.shift_indexes[start_time] = []
                self.shift_indexes[start_time].append(row)

            first_break = round(start_time) + 2
            second_break = round(end_time) - 2

            if first_break > 12:
                first_break -= 12
            if second_break < 0:
                second_break += 12
            elif second_break > 12:
                second_break -= 12

            # Write calculated break times to new excel file.
            style = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin')
            self.new_sheet.col(3).width = self.col_width
            self.new_sheet.write(row, 3, first_break, style)
            self.new_sheet.col(4).width = self.col_width
            self.new_sheet.write(row, 4, second_break, style)

        self.new_book.save(self.save_file_name)

    def calcLunchTimes(self):
        for start_time in self.shift_indexes:
            temp_count = 0
            count_modifier = 0
            num_employees = len(self.shift_indexes[start_time])
            #shift_count_midpoint = round(len(self.shift_indexes[start_time]) / 2)
            shift_count_quarter = round(num_employees / 3)
            for index in self.shift_indexes[start_time]:
                end_time = self.curr_sheet.cell(index, 2).value * 24

                # Lunch start_time starts 4 hours after start. Create base data that connects an 8oclock lunch to index 5.
                lunch_time = math.ceil(start_time) + 4

                '''                if num_employees > 4:
                    # Use base data to calculate the index for lunch start_time.
                    if start_time.is_integer():
                        lunch_index -= 1
                    else:
                        lunch_index -= 2

                    if temp_count >= shift_count_quarter:
                        count_modifier += 1
                        temp_count = 0
                    lunch_index += count_modifier'''

                # Use base data to calculate the index for lunch start_time.
                if start_time.is_integer():
                    lunch_index = self.open_hour[1] + (lunch_time - self.open_hour[0]) * 2 - 1
                else:
                    lunch_index = self.open_hour[1] + (lunch_time - self.open_hour[0]) * 2 - 2

                if num_employees > 4:
                    if temp_count >= shift_count_quarter:
                        count_modifier += 1
                        temp_count = 0
                    lunch_index += count_modifier
                else:
                    lunch_index += 1

                # Make sure cell border styling is correct according to where in hour lunch will occur.
                if lunch_index % 2:
                    style = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin')
                else:
                    style = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin')
                self.new_sheet.write(index, int(lunch_index), 'L', style)
                self.lunch_indexes.append((index, int(lunch_index)))

                # Create an employee shift object and add to appropriate shift list if not manager.
                if end_time - start_time <= 8:
                    new_employee_shift = EmployeeShift(index, int(lunch_index), start_time, end_time)
                    if start_time >= 14.5:
                        self.closing_shifts.append(new_employee_shift)
                    elif start_time >= 9.0:
                        self.mid_shifts.append(new_employee_shift)
                    elif start_time >= 6.5:
                        self.morning_shifts.append(new_employee_shift)
                    else:
                        self.open_shifts.append(new_employee_shift)
                    temp_count += 1
                    self.all_shifts = [self.open_shifts, self.morning_shifts, self.mid_shifts, self.closing_shifts]
        self.new_book.save(self.save_file_name)

    # Creates list of employees. Returns number of employees and list of employees working that day.
    def calcNumEmployees(self):
        employees = []
        for i in range(2, self.curr_sheet.nrows):
            name = self.curr_sheet.cell(i, 0)
            # Check if cell type not equal to 0 (0 represents empty).
            if name.value != '':
                employees.append(name.value)
        return len(employees), employees

    # Set cell to yellow.
    def setYellow(self, row, col):
        left_value = ''
        right_value = ''
        if (row, int(col)) in self.lunch_indexes:
            left_value = 'L'
        if (row, int(col+1)) in self.lunch_indexes:
            right_value = 'L'
        style = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; borders: left thin, top thin, bottom thin;')
        self.new_sheet.write(row, int(col), left_value, style)
        style = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; borders: right thin, top thin, bottom thin;')
        self.new_sheet.write(row, int(col+1), right_value, style)

        self.setReferenceCell(row, col, 'yellow')
        self.incrementReferenceCount(col)

        # Increment the employees CE count in reference matrix.
        self.reference_matrix[row-2][-1] += 1


    # Set cell to pink.
    def setPink(self, row, col):
        left_value = ''
        right_value = ''
        if (row, int(col)) in self.lunch_indexes:
            left_value = 'L'
        if (row, int(col + 1)) in self.lunch_indexes:
            right_value = 'L'
        style = xlwt.easyxf('pattern: pattern solid, fore_colour rose; borders: left thin, top thin, bottom thin;')
        self.new_sheet.write(row, int(col), left_value, style)
        style = xlwt.easyxf('pattern: pattern solid, fore_colour rose; borders: right thin, top thin, bottom thin;')
        self.new_sheet.write(row, int(col + 1), right_value, style)

        self.setReferenceCell(row, col, 'pink')

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
        self.createReferenceMatrix()

        # Category number 0=openshift, 1=morning, 2=mid, 3=closing
        category_num = 0
        for category in self.all_shifts:
            for shift in category:
                start_cell = self.translateHourToCell(shift.start_time)
                end_cell = self.translateHourToCell(shift.end_time)

                # Set product for first and last halfs if shift starts at half hour mark.
                if not shift.start_time.is_integer():
                    if start_cell > self.open_hour[1]:
                        style = xlwt.easyxf(
                            'pattern: pattern solid, fore_colour rose; borders: left thin, right thin, top thin, bottom thin;')
                        self.new_sheet.write(shift.row, int(start_cell)-1, '', style)
                    if end_cell <= 30:
                        style = xlwt.easyxf(
                            'pattern: pattern solid, fore_colour rose; borders: left thin, right thin, top thin, bottom thin;')
                        self.new_sheet.write(shift.row, int(end_cell)-2, "", style)
                        self.setReferenceCell(shift.row, int(end_cell)-2, 'pink')

                # If shift ends between 9:30 and 10 (starts between 1:30 and 2), give last hour CE
                if shift.start_time >= 13.5 and shift.start_time <= 14:
                    self.setYellow(shift.row, end_cell-4)
                '''
                if shift.start_time == 5.5:
                    self.setYellow(shift.row, start_cell+4)
                '''

                # Check if lunch given is at top or bottom of hour.
                lunch_top_of_hour = False
                if shift.lunch_index % 2:
                    lunch_top_of_hour = True

                # If shift before 5:30, set after lunch schedule as [P, Y, Y, P]
                if shift.start_time < 5.5:
                    curr_lunch_index = shift.lunch_index
                    curr_row = shift.row
                    self.setYellow(curr_row, curr_lunch_index-2)
                    self.setYellow(curr_row, curr_lunch_index+4)
                    self.setPink(curr_row, curr_lunch_index+6)

                # If 5:30 shift, set last half of shift to one of 2 shifts (add second shift later)
                if shift.start_time == 5.5:
                    self.setYellow(shift.row, shift.lunch_index + 1)
                    self.setYellow(shift.row, shift.lunch_index + 3)
                    self.setPink(shift.row, shift.lunch_index + 5)

                # Set lunch hour to pink for product. Also sets previous or after lunch hour to product.
                if lunch_top_of_hour:
                    self.setPink(shift.row, shift.lunch_index)
                    self.setPink(shift.row, shift.lunch_index + 2)
                else:
                    self.setPink(shift.row, shift.lunch_index-1)
                    if shift.lunch_index - 3 >= self.open_hour[1]:
                        self.setPink(shift.row, shift.lunch_index-3)

                # Set people that came in before or right after opening to CE first hour.
                if 7.0 <= shift.start_time <= 8.5:
                    self.setYellow(shift.row, self.open_hour[1])


            category_num += 1
        # Randomly fill in the rest of the cells until required CE count is met.
        for col in range(0, len(self.reference_matrix[self.num_employees + 1])):
            first_scheduled_hour = self.reference_matrix[self.num_employees + 1][col]
            last_scheduled_hour = self.reference_matrix[self.num_employees + 2][col]
            random_list = self.generateRandomCells(first_scheduled_hour, last_scheduled_hour)

            for row in random_list:
                if self.reference_matrix[row][col] == 'blank':
                    print("INside")
                    excel_col = self.translateHourToCell(col + 8)

                    if self.reference_matrix[self.num_employees][col] < self.recommended_ce[col]:
                        self.setYellow(row + 2, excel_col)
                        # self.reference_matrix[self.num_employees][col] += 1
                    else:
                        self.setPink(row + 2, excel_col)

        self.new_book.save(self.save_file_name)
        print("Here")

    def test_colorCells(self):
        last_time_recorded = 0
        time_count = 0

        for row in range(2, self.num_employees+2):
            start_time = self.curr_sheet.cell(row, 1).value * 24
            end_time = self.curr_sheet.cell(row, 2).value * 24
            if last_time_recorded != start_time:
                last_time_recorded = start_time
                time_count = 0
            else:
                time_count += 1
            index_count = 0
            index_modifier = 0
            shift_length = end_time - start_time
            if shift_length < 8.5 and shift_length > 5:
                hour_index = self.translateHourToCell(start_time)
                if hour_index < self.open_hour[1]:
                    hour_index = self.open_hour[1]

                '''
                if time_count < 4:
                    index_modifier = 0
                else:
                    index_modifier = 1
                '''

                for assignment in self.shift_assignments[self.shift_assignment_indices[start_time] + index_modifier]:
                    if assignment == 'yellow':
                        self.setYellow(row, hour_index + index_count)
                        index_count += 2
                    elif assignment == 'pink':
                        self.setPink(row, hour_index + index_count)
                        index_count += 2
                    time_count += 1

            elif shift_length < 5:
                shift_assignments = ['pink', 'pink', 'yellow', 'yellow']
                hour_index = self.translateHourToCell(start_time)
                for assignment in shift_assignments:
                    if assignment == 'yellow':
                        self.setYellow(row, hour_index + index_count)
                    elif assignment == 'pink':
                        self.setPink(row, hour_index + index_count)
                    time_count += 1
                    index_count += 2

        self.new_book.save(self.save_file_name)

    # Creates a matrix that will contain the color of each cell, to be used by randomized part of cell assignment.
    def createReferenceMatrix(self):
        # Both lists will keep track of the first and last scheduled rows for the current hour(column)
        # Last column keeps track of how many hours of CE an employee (row) has.
        first_scheduled_list = [-1]*13
        last_scheduled_list = [-1]*13

        for row in range(2, self.num_employees+2):
            reference_row = []
            for col in range(self.open_hour[1], 30, 2):

                # Extracting style information to obtain background colour of cell.
                xf = self.curr_sheet.cell_xf_index(row, col)
                xf_next = self.book.xf_list[xf]
                colour_index = xf_next.background.pattern_colour_index
                if colour_index == 22:
                    reference_row.append('gray')
                else:
                    reference_row.append('blank')

                reference_col = int((col-5) / 2)
                #print("Reference: ", first_scheduled_list[reference_col] == -1 and colour_index != 22)
                #print("Colour: ", colour_index)
                #print("Col: ", reference_col)

                ##### Minuses adjusting rows for both lists are due to the offset between the excel document and reference matrix.
                ##### Reference matrix's rows are 0 based index where excel is base 2.
                if row > self.last_manager_row:
                    if first_scheduled_list[reference_col] == -1 and colour_index != 22:
                        print("Setting First: ", row)
                        first_scheduled_list[reference_col] = row - 2
                    if first_scheduled_list[reference_col] != -1 and last_scheduled_list[reference_col] == -1 and \
                            last_scheduled_list[reference_col-1] < row and colour_index == 22:
                        print("Setting Last")
                        last_scheduled_list[reference_col] = row -2
                    if last_scheduled_list[reference_col] != -1 and colour_index != 22:
                        last_scheduled_list[reference_col] = -1

            reference_row.append(0)
            self.reference_matrix.append(reference_row)

        for col in range(0, len(last_scheduled_list)):
            if last_scheduled_list[col] == -1:
                last_scheduled_list[col] = self.num_employees

        self.reference_matrix.append([0]*13)
        self.reference_matrix.append(first_scheduled_list)
        self.reference_matrix.append(last_scheduled_list)

    def setReferenceCell(self, row, col, color):
        reference_col = int((col - 5) / 2)
        self.reference_matrix[row-2][reference_col] = color

    def incrementReferenceCount(self, col):
        reference_col = int((col - 5) / 2)
        print("Incrementing: ", reference_col)
        self.reference_matrix[self.num_employees][reference_col] += 1

    def getReferenceCell(self, row, col):
        reference_col = int((col - 5) / 2)
        return self.reference_matrix[row-2][reference_col]

    def generateRandomCells(self, low, high):
        cell_list = list(range(low, high+1))
        random.shuffle(cell_list)
        return cell_list

class EmployeeShift:

    def __init__(self, row, lunch_index, start_time, end_time):
        self.row = row
        self.lunch_index = lunch_index
        self.start_time = start_time
        self.end_time = end_time







                


