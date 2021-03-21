import os
from os.path import join, dirname, realpath
import pandas
from openpyxl import load_workbook
import datetime

if not os.path.exists("files"):
    os.mkdir("files")

path = join(dirname(realpath(__file__)), 'files/')
#########  ALGORITHM   ##########
# First read the file and save it into list
# Second convert the A column into String and datetime object and take diffrences of the consecutive rows
# If difference is greater than 1 then add a new row
# else don't do anythng
##################################
def add_blank_rows(filename):
    data = []
    data2 = []
    gap_rows = []
    workbook = load_workbook(filename=filename)
    sheet = workbook['Sheet1']
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    for i in range(0,len(data)-1):
        diff = data[i+1][0]-data[i][0]
        data2.append(list(data[i]))
        if int(str(diff).split(":")[2]) > 1:
            data2.append(list((" "," "," ")))
    for line in data2:
        #print(line)
        pass   
    df = pandas.DataFrame(data2)
    writer = pandas.ExcelWriter(path+filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Added_Blank_Space', index=False)
    writer.save()

print(os.listdir(path))
print(os.listdir(os.curdir))