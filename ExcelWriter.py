import xlwt
import math
import random
import time

class ExcelWriter:

    # new_book = New excel that will represent filled out schedule.
    # orig_book = The original excel file that all entries were copied from.
    def __init__(self, book, new_sheet, curr_sheet, new_book, styles, save_file_name, hour_ce_counts):
        self.book = book
        self.new_sheet = new_sheet
        self.curr_sheet = curr_sheet
        self.new_book = new_book
        self.num_employees = self.calcNumEmployees()[0]
        self.employees = self.calcNumEmployees()[1]
        self.half_employee_indexes = {}
        self.styles = styles
        self.open_hour = (curr_sheet.cell(1, 5).value, 5)
        self.shift_indexes = {}
        self.hour_shift_indexes = {}
        self.save_file_name = save_file_name
        self.col_width = 256 * 3
        self.last_manager_row = 0
        self.last_manager_row_reached = False
        self.row_ce_count = {}
        # Keeps track of the CE count and stores in corresponding buckets.
        self.ordered_ce_count = [[],[],[],[],[]]
        # Stores index of lunch for each row.
        self.row_lunch_column = []
        # Stores CE count of first and half ends of shift to be used for dispersment later.
        self.row_half_counts = []

        # Contains the indexes that should have an L (for lunch). Used for writing L's after coloring cells.
        self.lunch_indexes = []

        #Lists for sections of shifts.
        self.open_shifts = []
        self.morning_shifts = []
        self.mid_shifts = []
        self.closing_shifts = []
        self.half_shifts = []
        self.all_shifts = []

        #Recommended amount of CE members for each hour.
        self.recommended_ce = hour_ce_counts
        self.max_recommended_ce = [count + 3 for count in self.recommended_ce]

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
        # Off set of 2 since excel file starts with two information rows.
        for row in range(2, self.num_employees+2):
            # Calculating actual times since cell times given in percentage of a 24 hour day.
            start_time = self.curr_sheet.cell(row, 1).value * 24
            end_time = self.curr_sheet.cell(row, 2).value * 24

            # If manager (greater than 8 hours), record row. If row greater than certain number, don't change
            # because it might be merchant that was not moved to top of log.
            if end_time - start_time > 8 and not self.last_manager_row_reached:
                self.last_manager_row = row
            else:
                self.last_manager_row_reached = True

            # Keep track of indexes for each shift time.
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
                if index > self.last_manager_row:
                    end_time = self.curr_sheet.cell(index, 2).value * 24

                    # Lunch start_time starts 4 hours after start. Create base data that connects an 8oclock lunch to index 5.
                    lunch_time = math.ceil(start_time) + 4

                    # Use base data to calculate the index for lunch start_time.
                    if end_time - start_time >= 6:
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
                    if end_time - start_time < 6:
                        new_employee_shift = EmployeeShift(index, -1, start_time, end_time)
                        self.half_shifts.append(new_employee_shift)
                    elif end_time - start_time <= 10:
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

        self.all_shifts = [self.open_shifts, self.morning_shifts, self.mid_shifts, self.closing_shifts, self.half_shifts]
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
        ref_cel = int((col - 5) / 2)

        print("Setting Row: ", row)
        print("Setting Col: ", col)

        if self.reference_matrix[-3][ref_cel] < self.max_recommended_ce[ref_cel]:
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

            #print("Setting cell yellow: ", (row, col))
            self.setReferenceCell(row, col, 'yellow')
            self.incrementReferenceCount(col)

            # Add to first or second half of shift depending on location relative to lunch.
            # Offset row number to access lunch_column and half_counts since index 0 these lists refers to rows after
            # last manager row.
            row_offset = row-self.last_manager_row-1
            if ref_cel < self.row_lunch_column[row_offset]:
                #print("Test: ", self.row_half_counts[row-2][0])
                self.row_half_counts[row_offset][0] += 1
            elif ref_cel > self.row_lunch_column[row_offset]:
                self.row_half_counts[row_offset][1] += 1

            # Increment the employees CE count in reference matrix.
            self.reference_matrix[row-2][-1] += 1

            #self.row_ce_count[row] += 1


    # Set cell to pink.
    def setPink(self, row, col, fromYellow):
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

        # Decrement the employees CE count in reference matrix.
        if fromYellow:
            self.reference_matrix[row - 2][-1] -= 1
            self.decrementReferenceCount(col)

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

        # Category number 0=openshift, 1=morning, 2=mid, 3=closing 4=half shifts
        for category in self.all_shifts:
            for shift in category:
                print("Inserting: ", shift.row)
                start_cell = self.translateHourToCell(shift.start_time)
                end_cell = self.translateHourToCell(shift.end_time)
                new_ce_count = 0

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
                '''
                # If shift ends between 9:30 and 10 (starts between 1:30 and 2), give last hour CE
                if shift.start_time >= 13 and shift.start_time <= 14:
                    self.setYellow(shift.row, end_cell-4)
                    new_ce_count += 1
                '''
                if category != self.half_shifts:
                    # Check if lunch given is at top or bottom of hour.
                    lunch_top_of_hour = False
                    if shift.lunch_index % 2:
                        lunch_top_of_hour = True

                    '''
                    # If shift before 5:30, set after lunch schedule as [P, Y, Y, P]
                    if shift.start_time < 5.5 and shift.start_time >= 5:
                        curr_lunch_index = shift.lunch_index
                        curr_row = shift.row
                        self.setYellow(curr_row, curr_lunch_index-2)
                        self.setYellow(curr_row, curr_lunch_index+4)
                        new_ce_count += 2
                        self.setPink(curr_row, curr_lunch_index+6, False)
    
                    # If 5:30 shift, set last half of shift to one of 2 shifts (add second shift later)
                    if shift.start_time == 5.5:
                        self.setYellow(shift.row, shift.lunch_index + 1)
                        self.setYellow(shift.row, shift.lunch_index + 3)
                        new_ce_count += 2
                        self.setPink(shift.row, shift.lunch_index + 5, False)
                    '''
                    # Set lunch hour to pink for product. Also sets previous or after lunch hour to product.
                    lunch_col = 0
                    if lunch_top_of_hour:
                        self.setPink(shift.row, shift.lunch_index, False)
                        #self.setPink(shift.row, shift.lunch_index + 2, False)
                        self.setReferenceCell(shift.row, shift.lunch_index, 'top_lunch')
                        lunch_col = int((shift.lunch_index - 5) / 2)
                    else:
                        self.setPink(shift.row, shift.lunch_index-1, False)
                        self.setReferenceCell(shift.row, shift.lunch_index-1, 'bottom_lunch')
                        lunch_col = int(((shift.lunch_index - 1) - 5) / 2)
                        #if shift.lunch_index - 3 >= self.open_hour[1]:
                            #self.setPink(shift.row, shift.lunch_index-3, False)

                    self.row_lunch_column.append(lunch_col)

                else:
                    self.row_lunch_column.append(-1)
                self.row_half_counts.append([0,0])

                '''
                # Set people that came in before or right after opening to CE first hour.
                if 7.0 <= shift.start_time <= 8.5:
                    self.setYellow(shift.row, self.open_hour[1])
                    new_ce_count += 1
                '''
                # Move row to corresponding bucket.
                for bucket in self.ordered_ce_count:
                    if shift.row in bucket:
                        bucket.remove(shift.row)
                        index = self.ordered_ce_count.index(bucket)
                        self.ordered_ce_count[index+new_ce_count].append(shift.row)
                        break

        # Randomly fill in the rest of the cells until required CE count is met.
        # This list will maintain all rows that were set to CE in previous column. Will get replaced every col.
        # Will be reset later, but needed now to enter loop.
        prev_col_ce = []

        for col in range(0, len(self.reference_matrix[self.num_employees + 1])):
            # First and last scheduled hour variables measured from reference matrix, not excel (starts at 0, not 2)
            first_scheduled_hour = self.reference_matrix[self.num_employees + 1][col]
            last_scheduled_hour = self.reference_matrix[self.num_employees + 2][col]

            number_of_passes = 0
            while number_of_passes <= 2:
                for bucket in self.ordered_ce_count:
                    random.shuffle(bucket)
                    for row in bucket:
                        if self.reference_matrix[row][col] == 'blank':
                            excel_col = self.translateHourToCell(col + 8)

                            # For first pass. If blank and the previous column was not just set to CE (also following placing rules)
                            # then set cell to yellow, and add to prev_col_ce.
                            if self.reference_matrix[self.num_employees][col] < self.recommended_ce[col] and self.shouldPlaceCe(row, col):
                                # Check for first pass
                                if number_of_passes == 0 and row not in prev_col_ce:
                                    self.setYellow(row + 2, excel_col)
                                else:
                                    self.setYellow(row + 2, excel_col)

                                # Change 'Bucket' that row is in (move into next num_ce bucket)
                                bucket.remove(row)
                                current_bucket_index = self.ordered_ce_count.index(bucket)
                                self.ordered_ce_count[current_bucket_index+1].append(row)
                number_of_passes += 1

            # After setting all possible rows to CE, add these rows to the prev_col_ce list.
            prev_col_ce = []
            for row in range(first_scheduled_hour, last_scheduled_hour):
                if self.reference_matrix[row][col] == 'yellow':
                    prev_col_ce.append(row)
                else:
                    self.setPink(row+2, excel_col, False)


        # Even out hour distribution by going through row_half_counts and moving around peoples hours.
        oversized_distribution_rows_left = []
        oversized_distribution_rows_right = []
        '''
        for category in self.all_shifts:
            for shifts in category:
                pass
        '''
        # overpopulated right and left sides. Adds last_manager_row to make up for offset from ignoring managers.
        for row in range(0, len(self.row_half_counts)):
            if (self.row_half_counts[row][1] - self.row_half_counts[row][0]) > 1:
                oversized_distribution_rows_right.append(row+self.last_manager_row-1)
            elif (self.row_half_counts[row][0] - self.row_half_counts[row][1]) > 1:
                oversized_distribution_rows_left.append(row+self.last_manager_row-1)



        # Go through and try and even out the overpopulated sides.
        for right_row in oversized_distribution_rows_right:
            for left_row in oversized_distribution_rows_left:
                right_lunch = self.row_lunch_column[right_row-self.last_manager_row]
                left_lunch = self.row_lunch_column[left_row-self.last_manager_row]
                if 0 <= left_lunch - right_lunch <= 4:
                    for cols in range(right_lunch+1, right_lunch+5):

                        if cols <= len(self.reference_matrix[right_row])-1 and self.shouldPlaceCe(right_row, cols)\
                                and self.reference_matrix[left_row][cols] == 'yellow':
                            excel_col = 2 * cols + 5
                            self.setPink(left_row+2, excel_col, True)
                            self.setYellow(right_row+2, excel_col)

        self.new_book.save(self.save_file_name)
        print("Ending coloring cells")

    # Creates a matrix that will contain the color of each cell, to be used by randomized part of cell assignment.
    def createReferenceMatrix(self):
        # Both lists will keep track of the first and last scheduled rows for the current hour(column)
        # Last column keeps track of how many hours of CE an employee (row) has.
        first_scheduled_list = [-1]*13
        last_scheduled_list = [-1]*13

        for row in range(2, self.num_employees+2):
            reference_row = []
            if row > self.last_manager_row:
                self.ordered_ce_count[0].append(row-2)
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

                ##### Minuses adjusting rows for both lists are due to the offset between the excel document and reference matrix.
                ##### Reference matrix's rows are 0 based index where excel is base 2.
                if row > self.last_manager_row:
                    if first_scheduled_list[reference_col] == -1 and colour_index != 22:
                        first_scheduled_list[reference_col] = row - 2
                    if first_scheduled_list[reference_col] != -1 and last_scheduled_list[reference_col] == -1 and \
                            last_scheduled_list[reference_col-1] < row and colour_index == 22:
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
        if (reference_col > 13):
            print("Wrong: ", row, col)
        self.reference_matrix[row-2][reference_col] = color

    def incrementReferenceCount(self, col):
        reference_col = int((col - 5) / 2)
        self.reference_matrix[self.num_employees][reference_col] += 1

    def decrementReferenceCount(self, col):
        reference_col = int((col - 5) / 2)
        self.reference_matrix[self.num_employees][reference_col] -= 1

    def getReferenceCell(self, row, col):
        reference_col = int((col - 5) / 2)
        return self.reference_matrix[row-2][reference_col]

    def generateRandomCells(self, low, high):
        cell_list = list(range(low, high+1))
        random.shuffle(cell_list)
        return cell_list


    # Method checks surrounding cells and decides if CE cell should be placed in current column.
    def shouldPlaceCe(self, row, col):
        #Create random variable that decides if there should be two hours in a row.
        answer = True
        if col <= len(self.reference_matrix[row]) - 1:
            if self.reference_matrix[row][col] == 'yellow':
                answer = False
            if self.reference_matrix[row][col] == 'top_lunch' or self.reference_matrix[row][col] == 'bottom_lunch':
                answer = False
            if self.reference_matrix[row][col] == 'gray':
                answer = False
                '''
            if self.reference_matrix[row][-1] > 2:
                answer = False
                '''
        try:
            if self.reference_matrix[row][col-1] == 'yellow' and self.reference_matrix[row][col-2] == 'yellow':
                answer = False
            if self.reference_matrix[row][col+1] == 'yellow' and self.reference_matrix[row][col+2] == 'yellow':
                answer = False
            if self.reference_matrix[row][col-1] == 'yellow' and self.reference_matrix[row][col+1] == 'yellow':
                answer = False

        except:
            answer = True
        return answer

    # Method checks to see if there are four product rows in a row to place a new CE row.
    def threePinkInRow(self, row, col):
        answer = False
        placing_col = -1
        '''
        if self.reference_matrix[row][col] == 'pink':
            if self.reference_matrix[row][col-1] == 'pink' and self.self.reference_matrix[row][col-2] == 'pink':
                answer = True
                if self.reference_matrix[row][col-3] == 'lunch': <---- check top/bottom lunch
                    placing_col = col-1
                else:
                    placing_col = col-2
            elif self.reference_matrix[row][col-1] == 'pink' and self.reference_matrix[row][col+1] == 'pink':
                answer = True
                placing
            elif self.reference_matrix[row][col+1] == 'pink' and self.reference_matrix[row][col+2] == 'pink':
                answer = True
        '''
        try:
            if self.reference_matrix[row][-1] < 5:
                #print("In First")
                if self.reference_matrix[row][col] == 'pink' and 0 < col < len(self.reference_matrix[0])-2: #Previously 0 and -4
                    #print("In Second")
                    pt_logic = self.three_in_row_helper_top(row, col)
                    if pt_logic:
                        if (self.reference_matrix[row][col-1] != 'yellow' and self.reference_matrix[row][col+1] != 'yellow' and
                            self.reference_matrix[row][col] != 'top_lunch' and self.reference_matrix[row][col] != 'bottom_lunch'):
                            answer = True
                            placing_col = col

                        '''
                        pt_logic = ((self.reference_matrix[row][col+1] == 'pink' or self.reference_matrix[row][col+1] == 'top_lunch'
                                or self.reference_matrix[row][col +1] == 'bottom_lunch') and (self.reference_matrix[row][col +2] == 'pink' or
                                self.reference_matrix[row][col +2] == 'top_lunch' or self.reference_matrix[row][col +2] == 'bottom_lunch'))
                                
                        print("In Third")
                        answer = True
                        if self.reference_matrix[row][col-1] == 'top_lunch':
                            if self.reference_matrix[row][col+2] == 'yellow':
                                placing_col = col
                            else:
                                placing_col = col + 1
                        elif self.reference_matrix[row][col-1] == 'bottom_lunch':
                            placing_col = col
                        elif self.reference_matrix[row][col+3] == 'top_lunch':
                            placing_col = col + 2
                        elif self.reference_matrix[row][col+3] == 'bottom_lunch':
                            placing_col = col + 1
                        else:
                            random_modifier = random.randint(0,1)
                            if placing_col == 12:
                                random_modifier = 0
                            placing_col = col + random_modifier
                        '''
        except():
            print("Index was out of range")
        print("Answer: ", answer)
        return (answer, placing_col)

    def three_in_row_helper_top(self, row, col):
        if self.three_in_row_helper_second(row,col-1) and self.three_in_row_helper_second(row, col+1):
            return True
        elif self.three_in_row_helper_second(row,col+1) and self.three_in_row_helper_second(row, col+2):
            return True
        elif self.three_in_row_helper_second(row,col-1) and self.three_in_row_helper_second(row, col-2):
            return True
        else:
            return False

    def three_in_row_helper_second(self, row, col):
        return self.ref_help(row,col) == 'pink' or self.ref_help(row, col) == 'top_lunch' or self.ref_help(row,col) == 'bottom_lunch'

    # Helper function just returns the value of the reference matrix cell at row,col. For readability.
    def ref_help(self, row, col):
        return self.reference_matrix[row][col]



class EmployeeShift:

    def __init__(self, row, lunch_index, start_time, end_time):
        self.row = row
        self.lunch_index = lunch_index
        self.start_time = start_time
        self.end_time = end_time







                


