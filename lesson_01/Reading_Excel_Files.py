#!/usr/bin/env python
"""
Your task is as follows:
- read the provided Excel file
- find and return the min, max and average values for the COAST region
- find and return the time value for the min and max entries
- the time values should be returned as Python tuples

Please see the test function for the expected return format
"""

import xlrd
from zipfile import ZipFile
datafile = "../data/2013_ERCOT_Hourly_Load_Data.xls"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)

    ### example on how you can get the data
    sheet_data = [[sheet.cell_value(r, col) for col in range(sheet.ncols)] for r in range(sheet.nrows)]


    coast_list = sheet.col_values(1, start_rowx=1)

    avg_coast = sum(coast_list) / len(coast_list)

    max_coast_value = max(coast_list)
    max_coast_index = coast_list.index(max_coast_value)+1
    max_coast_time = xlrd.xldate_as_tuple(sheet.cell_value(max_coast_index,0),0)

    min_coast_value = min(coast_list)
    min_coast_index = coast_list.index(min_coast_value)+1
    min_coast_time = xlrd.xldate_as_tuple(sheet.cell_value(min_coast_index,0),0)

    print ('avg coast value: ' + str(avg_coast))

    print ('max coast index: ' + str(max_coast_index))
    print ('max coast value: ' + str(max_coast_value))
    print ('max coast time:  ' + str(max_coast_time))

    print ('min coast index: ' + str(min_coast_index))
    print ('min coast value: ' + str(min_coast_value))
    print ('min coast time: ' + str(min_coast_time))
    

    ### other useful methods:
    # print ("\nROWS, COLUMNS, and CELLS:")

    # # print (max(sheet.col(1))
    # print ("Number of rows in the sheet:"),
    # print (sheet.nrows)
    # print ("Type of data in cell (row 3, col 2):"), 
    # print (sheet.cell_type(3, 2))
    # print ("Value in cell (row 3, col 2):"), 
    # print (sheet.cell_value(3, 2))
    # print ("Get a slice of values in column 3, from rows 1-3:")
    # print (sheet.col_values(3, start_rowx=1, end_rowx=4))

    # print ("\nDATES:")
    # print ("Type of data in cell (row 1, col 0):"), 
    # print (sheet.cell_type(1, 0))
    # exceltime = sheet.cell_va,olue(1, 0)
    # print ("Time in Excel format:"),
    # print (exceltime)
    # print ("Convert time to a Python datetime tuple, from the Excel float:"),
    # print (xlrd.xldate_as_tuple(exceltime, 0))
    
    
    data = {
            'maxtime': max_coast_time,
            'maxvalue': max_coast_value,
            'mintime': min_coast_time,
            'minvalue': min_coast_value,
            'avgcoast': avg_coast
    }
    return data


def test():
    # open_zip(datafile)
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)


test()