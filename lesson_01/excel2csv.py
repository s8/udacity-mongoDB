# -*- coding: utf-8 -*-
# Find the time and value of max load for each of the regions
# COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
# and write the result out in a csv file, using pipe character | as the delimiter.
# An example output can be seen in the "example.csv" file.

import xlrd
import os
import csv
from zipfile import ZipFile

datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    data = None
    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)

    stations = [s.value for s in sheet.row(0)[1:-1]]    
    times  = [xlrd.xldate_as_tuple(t.value,0) for t in sheet.col(0)[1:]]
    loads = [[t.value for t in sheet.col(c)[1:]] for c in range(sheet.ncols)[1:]]

    data = []

    for i, s in enumerate(stations):
        load = max(loads[i])
        load_i = loads[i].index(load)
        time = times[load_i]

        data.append({
            s:{
            'Max Load' : load,
            'Year' : time[0],
            'Month' : time[1],
            'Day' : time[2],
            'Hour' : time[3]
            }})

    return data

def save_file(data, filename):
    stations = [list(d.keys())[0] for d in data]
    years = [list(d.keys())[0] for d in data]

    with open(filename, 'wt', encoding='ascii') as f:
        csvwriter = csv.writer(f, delimiter='|')
        for i, s in enumerate(stations):
            d = data[i][s]
            csvwriter.writerow([s, d['Year'],
                d['Month'],d['Day'],d['Hour'],d['Max Load']])

    
def test():
    open_zip(datafile)
    data = parse_file(datafile)
    save_file(data, outfile)

    number_of_rows = 0
    stations = []

    ans = {'FAR_WEST': {'Max Load': '2281.2722140000024',
                        'Year': '2013',
                        'Month': '6',
                        'Day': '26',
                        'Hour': '17'}}
    correct_stations = ['COAST', 'EAST', 'FAR_WEST', 'NORTH',
                        'NORTH_C', 'SOUTHERN', 'SOUTH_C', 'WEST']
    fields = ['Year', 'Month', 'Day', 'Hour', 'Max Load']

    with open(outfile) as of:
        csvfile = csv.DictReader(of, delimiter="|")
        for line in csvfile:
            station = line['Station']
            if station == 'FAR_WEST':
                for field in fields:
                    # Check if 'Max Load' is within .1 of answer
                    if field == 'Max Load':
                        max_answer = round(float(ans[station][field]), 1)
                        max_line = round(float(line[field]), 1)
                        assert max_answer == max_line

                    # Otherwise check for equality
                    else:
                        assert ans[station][field] == line[field]

            number_of_rows += 1
            stations.append(station)

        # Output should be 8 lines not including header
        assert number_of_rows == 8

        # Check Station Names
        assert set(stations) == set(correct_stations)

        
if __name__ == "__main__":
    test()
