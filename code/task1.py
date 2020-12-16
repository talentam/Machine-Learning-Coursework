import datetime

import openpyxl
from openpyxl import load_workbook
import numpy as np


# read worksheet and find out the overlapped year
def read_worksheet(sheetName):
    data_list = []
    start_year = 3000
    end_year = 1000
    for row in range(2, sheetName.max_row + 1):
        column4 = sheetName.cell(row=row, column=4).value
        column5 = sheetName.cell(row=row, column=5).value
        column6 = sheetName.cell(row=row, column=6).value
        column7 = sheetName.cell(row=row, column=7).value
        temporary_list = [column4, column5, column6, column7]
        data_list.append(temporary_list)

        if column5.year < start_year:
            start_year = column5.year
        elif column5.year > end_year:
            end_year = column5.year

    return data_list, start_year, end_year


# read workbook
def read_workbook(path):
    wb = load_workbook(path)
    Chla = wb['CHLA ']
    Temp = wb['TEMPERATURE']
    TotalP = wb['Total P ']
    Chla_data, start_year1, end_year1 = read_worksheet(Chla)
    Temp_data, start_year2, end_year2 = read_worksheet(Temp)
    TotalP_data, start_year3, end_year3 = read_worksheet(TotalP)
    start_year_max = max(start_year1, start_year2, start_year3)
    end_year_min = min(end_year1, end_year2, end_year3)

    return Chla_data, Temp_data, TotalP_data, start_year_max, end_year_min


# remove the rows that some cell is empty or the year is not overlapped or month exceed Apr-Nov
def data_cleaning(data, year_start, year_end, month_start, month_end):
    remove_index = []
    for i, value in enumerate(data):
        length = len(value)
        if value[length-1] is None or value[length-2] is None or value[length-3].year < year_start \
                or value[length-3].year > year_end or value[length-4] == 2 \
                or value[length-3].month < month_start or value[length-3].month > month_end:
            remove_index.append(i)

    remove_index.reverse()
    for index in remove_index:
        data.pop(index)


def find_data_needed_completed(Chla, Temp, TotalP):
    temp_complete_index = []
    totalP_complete_index = []
    for i, day1 in enumerate(Chla):

        # for any row in Chla, if there is no corresponding same day and depth, it need to be complete
        counter = 0
        for day2 in Temp:
            # if day1[1] not in day2 or day1[2] not in day2:
            if day1[1] == day2[1] and day1[2] == day2[2]:
                counter = counter + 1
                break
        if counter == 0:
            temp_complete_index.append(i)

        counter = 0
        for day2 in TotalP:
            # if day1[1] not in day2 or day1[2] not in day2:
            if day1[1] == day2[1] and day1[2] == day2[2]:
                counter = counter + 1
                break
        if counter == 0:
            totalP_complete_index.append(i)
    return temp_complete_index, totalP_complete_index


# read the workbook and find overlapped years
print('[INFO] reading workbook')
Chla_list, Temp_list, TotalP_list, starting_year, ending_year = read_workbook('./lake_data/China lake.xlsx')

# the rows contains empty cell will be removed
# Also, the year outside the overlapped period will be removed
# For Chal table, only 5-10 months will be kept
# For other tables, only 4-11 months will be kept
print('[INFO] cleaning data')
starting_month, ending_month = 5, 10
data_cleaning(Chla_list, starting_year, ending_year, starting_month, ending_month)

starting_month, ending_month = 4, 11
data_cleaning(Temp_list, starting_year, ending_year, starting_month, ending_month)
data_cleaning(TotalP_list, starting_year, ending_year, starting_month, ending_month)

# find the data which need to be completed
Temp_complete, TotalP_complete = find_data_needed_completed(Chla_list, Temp_list, TotalP_list)




