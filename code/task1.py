import datetime

import openpyxl
from openpyxl import load_workbook, Workbook
import numpy as np
from collections import Counter


# read worksheet and find out the overlapped year
def read_worksheet(sheetName):
    # store the all the data in the list
    data_list = []
    # this list used for finding the depth that appears most
    depth_list = []
    start_year = 3000
    end_year = 1000
    for row in range(2, sheetName.max_row + 1):
        column4 = sheetName.cell(row=row, column=4).value
        column5 = sheetName.cell(row=row, column=5).value
        column6 = sheetName.cell(row=row, column=6).value
        column7 = sheetName.cell(row=row, column=7).value
        temporary_list = [column4, column5, column6, column7]
        data_list.append(temporary_list)
        depth_list.append(column6)

        if column5.year < start_year:
            start_year = column5.year
        elif column5.year > end_year:
            end_year = column5.year

    return data_list, depth_list, start_year, end_year


# read workbook
def read_workbook(path):
    workbook = load_workbook(path)
    Chla_ws = workbook['CHLA ']
    Temp_ws = workbook['TEMPERATURE']
    TotalP_ws = workbook['Total P ']
    Chla_data, Chla_depth, start_year1, end_year1 = read_worksheet(Chla_ws)
    Temp_data, Temp_depth, start_year2, end_year2 = read_worksheet(Temp_ws)
    TotalP_data, TotalP_depth, start_year3, end_year3 = read_worksheet(TotalP_ws)

    # find out the start year ,end year
    start_year_max = max(start_year1, start_year2, start_year3)
    end_year_min = min(end_year1, end_year2, end_year3)

    # find out the majority depth
    depth_total = Chla_depth + Temp_depth + TotalP_depth
    max_depth = Counter(depth_total).most_common(1)[0][0]

    return Chla_data, Temp_data, TotalP_data, start_year_max, end_year_min, max_depth


# remove the rows that some cell is empty or the year is not overlapped or month exceed Apr-Nov
def data_cleaning(data, year_start, year_end, month_start, month_end, max_depth):
    remove_index = []
    for i, value in enumerate(data):
        length = len(value)
        if value[length-1] is None or value[length-2] is None or value[length-3].year < year_start \
                or value[length-3].year > year_end or value[length-4] == 2 \
                or value[length-3].month < month_start or value[length-3].month > month_end \
                or value[length-2] != max_depth:
            remove_index.append(i)

    remove_index.reverse()
    for index in remove_index:
        data.pop(index)


def initializeEmptyList(start_year, end_year, start_month, end_month):
    rows = end_year - start_year + 1
    columns = end_month - start_month + 1
    empty_list = [[[] for i in range(columns)] for j in range(rows)]
    return empty_list


def matchData(original_list, complete_list):
    year_num = ending_year - starting_year + 1
    month_num = ending_month - starting_month + 1

    for row in original_list:
        year = row[1].year - starting_year
        month = row[1].month - starting_month
        complete_list[year][month].append(row[3])

    for i in range(year_num):
        for j in range(month_num):
            if len(complete_list[i][j]) == 0:
                complete_list[i][j] = [0]
            else:
                complete_list[i][j] = [np.mean(complete_list[i][j])]

    return complete_list


# count zero numbers
def countZero(input_list):
    zero_num = 0
    for i in input_list:
        if i[0] == 0:
            zero_num = zero_num + 1
    return zero_num


def meanCalculation(input_list):
    month_num = ending_month - starting_month + 1
    for year in input_list:
        # print(year)
        # count the zero number
        zero_num = countZero(year)
        while 0 < zero_num < 4:
            # check condition of 101
            for i in range(0, month_num-2):
                if year[i][0] != 0 and year[i+1][0] == 0 and year[i+2][0] != 0:
                    year[i + 1][0] = (year[i][0] + year[i + 2][0])/2
                    zero_num = zero_num - 1
                    if zero_num == 0:
                        break

            # check condition of 110 or 011
            for i in range(0, month_num - 2):
                # 110
                if year[i][0] != 0 and year[i + 1][0] != 0 and year[i + 2][0] == 0:
                    year[i + 2][0] = 2*year[i + 1][0] - year[i][0]
                    zero_num = zero_num - 1
                    break
                # 011
                elif year[i][0] == 0 and year[i + 1][0] != 0 and year[i + 2][0] != 0:
                    year[i][0] = 2 * year[i + 1][0] - year[i+2][0]
                    zero_num = zero_num - 1
                    break

    # remove year without data
    # remove_index = []
    # for i, year in enumerate(input_list):
    #     if countZero(year) > 0:
    #         remove_index.append(i)
    #
    # remove_index.reverse()
    # for i in remove_index:
    #     input_list.pop(i)

    return input_list


def outputTable(list1, list2, list3, info_list, table_name):
    # wb = Workbook()
    sheet = wb.create_sheet(table_name)
    sheet.append(['MIDAS', 'LAKE', 'Town(s)', 'STATION', 'Date', 'DEPTH', 'CHLA (mg/L)', 'TEMPERATURE (Centrigrade)', 'Total P (mg/L)'])
    for i in range(0, len(list1) * len(list1[0])):
        if list1[i // 6][i % 6][0] != 0 and list2[i // 6][i % 6][0] != 0 and list3[i // 6][i % 6][0] != 0:
            # append normal information
            row = ['5448', 'China Lake', 'China, Vassalboro']
            # append station
            row.append(info_list[0][0])
            # append year/month
            year = i // 6 + starting_year
            month = i % 6 + starting_month
            year_month = str(year)+'/'+str(month)
            row.append(year_month)
            # append depth
            row.append(info_list[0][2])
            # append CHLA
            row.append(list1[i // 6][i % 6][0])
            # append TEMPERATURE
            row.append(list2[i // 6][i % 6][0])
            # append TotalP
            row.append(list3[i // 6][i % 6][0])

            sheet.append(row)

    # wb.save("./lake_data/completeChinaLake.xlsx")


# read the workbook and find overlapped years
print('[INFO] reading workbook')
Chla_list, Temp_list, TotalP_list, starting_year, ending_year, depth = read_workbook('./lake_data/China lake.xlsx')

# the rows contains empty cell will be removed
# Also, the year outside the overlapped period will be removed
# For Chal table, only 5-10 months will be kept
# For other tables, only 4-11 months will be kept
print('[INFO] cleaning data')
starting_month, ending_month = 5, 10
data_cleaning(Chla_list, starting_year, ending_year, starting_month, ending_month, depth)
data_cleaning(Temp_list, starting_year, ending_year, starting_month, ending_month, depth)
data_cleaning(TotalP_list, starting_year, ending_year, starting_month, ending_month, depth)

# find the data which need to be completed
# Temp_complete, TotalP_complete = find_data_needed_completed(Chla_list, Temp_list, TotalP_list)

# initialize empty complete list
Chla = initializeEmptyList(starting_year, ending_year, starting_month, ending_month)
Temp = initializeEmptyList(starting_year, ending_year, starting_month, ending_month)
TotalP = initializeEmptyList(starting_year, ending_year, starting_month, ending_month)

# match data which does not need mean calculation
print('[INFO] matching data')
Chla = matchData(Chla_list, Chla)
Temp = matchData(Temp_list, Temp)
TotalP = matchData(TotalP_list, TotalP)

# method 1: mean value calculation
print('[INFO] mean value calculation')
Chla = meanCalculation(Chla)
Temp = meanCalculation(Temp)
TotalP = meanCalculation(TotalP)

# method 2:

# output table to excel
wb = Workbook()

outputTable(Chla, Temp, TotalP, Chla_list, 'method 1')
outputTable(Chla, Temp, TotalP, Chla_list, 'method 2')

wb.save("./lake_data/completeChinaLake.xlsx")














