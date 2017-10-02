import xlrd
import os
import csv
from zipfile import ZipFile

datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()
'''
[text:u'Hour_End', text:u'COAST', text:u'EAST', text:u'FAR_WEST', text:u'NORTH', 
text:u'NORTH_C', text:u'SOUTHERN', text:u'SOUTH_C', text:u'WEST', text:u'ERCOT']
'''
def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    data = []

    max_coast =     max(sheet.col_values(1, start_rowx=1, end_rowx=sheet.nrows))
    max_east =      max(sheet.col_values(2, start_rowx=1, end_rowx=sheet.nrows))
    max_far_west =  max(sheet.col_values(3, start_rowx=1, end_rowx=sheet.nrows))
    max_north =     max(sheet.col_values(4, start_rowx=1, end_rowx=sheet.nrows))
    max_north_c =   max(sheet.col_values(5, start_rowx=1, end_rowx=sheet.nrows))
    max_southern =  max(sheet.col_values(6, start_rowx=1, end_rowx=sheet.nrows))
    max_south_c =   max(sheet.col_values(7, start_rowx=1, end_rowx=sheet.nrows))
    max_west =      max(sheet.col_values(8, start_rowx=1, end_rowx=sheet.nrows))
    
    time_column = sheet.col_values(0, start_rowx=1, end_rowx=sheet.nrows)
    for i in range(1,9):
        region_name = sheet.cell_value(0, i)
        region_column = sheet.col_values(i, start_rowx=1, end_rowx=sheet.nrows)
        for time,value in zip(time_column, region_column):
            time_tuple = xlrd.xldate_as_tuple(time,0)
            if float(value) == float(max_coast):
                data.append([region_name, time_tuple, max_coast])
            elif float(value) == float(max_east):
                data.append([region_name, time_tuple, max_east])
            elif float(value) == float(max_far_west):
                data.append([region_name, time_tuple, max_far_west])
            elif float(value) == float(max_north):
                data.append([region_name, time_tuple, max_north])
            elif float(value) == float(max_north_c):
                data.append([region_name, time_tuple, max_north_c])
            elif float(value) == float(max_southern):
                data.append([region_name, time_tuple, max_southern])
            elif float(value) == float(max_south_c):
                data.append([region_name, time_tuple, max_south_c])
            elif float(value) == float(max_west):
                data.append([region_name, time_tuple, max_west])
    return data

def save_file(data, filename):
    f = open(filename,'wb')
    w = csv.writer(f)
    
    w.writerow(["Station|Year|Month|Day|Hour|Max Load"])
    
    for i in range(len(data)):
        region_name =   str(data[i][0])
        year =          str(data[i][1][0])
        month =         str(data[i][1][1])
        day =           str(data[i][1][2])
        hour =          str(data[i][1][3])
        region_value =  str(data[i][2])
        
        w.writerow([region_name+"|"+year+"|"+month+"|"+day+"|"+hour+"|"+region_value])
    f.close()
    
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
            print line
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
