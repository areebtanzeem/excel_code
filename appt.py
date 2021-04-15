import os
import sys
from os.path import join, dirname, realpath
import pandas
from openpyxl import load_workbook
import datetime
from shutil import copyfile
import numpy as np
import operator
import math

if not os.path.exists("files"):
    os.mkdir("files")

path = join(dirname(realpath(__file__)), 'files/')
#########  ALGORITHM   ##########
# First read the file and save it into list
# Second convert the A column into String and datetime object and take diffrences of the consecutive rows
# If difference is greater than 1 then add a new row
# else don't do anythng
##################################


def main(filename):
    df = pandas.read_excel(filename, names=['one', 'two', 'three'],header=None)
    writer1 = pandas.ExcelWriter(path+"temp_"+filename, engine='xlsxwriter')
    df.to_excel(writer1, sheet_name='Added_Blank_Space', index=False)
    writer1.save()

    # For B column average and for C column sun
    #print("Duplicate entries for file: ", filename)
    #ids = df["one"]
    #print(pandas.concat(g for _, g in df.groupby("one") if len(g) > 1))

    df2 = df.groupby(['one'], as_index=False).agg(
        {'two': 'mean', 'three': 'sum'})

    # For B column average and for C column sun Ends Here

    #df.drop_duplicates(subset=['one'], keep=False, inplace=True)
    df2['difference'] = df2.one.diff(1)
    list = df2.values.tolist()
    list2 = []
    for i in list:
        if i[3] > datetime.timedelta(seconds=1):
            list2.append(["", "", ""])
            list2.append([i[0], i[1], i[2]])
        else:
            list2.append([i[0], i[1], i[2]])

    df3 = pandas.DataFrame(list2, columns=['one', 'two', 'three'])
    writer = pandas.ExcelWriter(path+filename, engine='xlsxwriter')
    df3.to_excel(writer, sheet_name='Added_Blank_Space', index=False)
    writer.save()
    print(df3)
    # print(df.dtypes)


def add_blank_rows(filename):
    data = []
    data2 = []
    gap_rows = []
    duplicate_index = []
    workbook = load_workbook(filename=filename)
    sheet = workbook['Sheet1']
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    # for i in range(0,len(data)-1):
    #    diff = data[i+1][0]-data[i][0]
    #    print(data[i+1][0])
    #    print(diff)
    #    data2.append(list(data[i]))
    #    if int(float(str(diff).split(":")[2])) > 1:
    #        data2.append(list(("","","")))
    #    elif int(float(str(diff).split(":")[2])) == 0:
    #        duplicate_index.append(i)
    #        duplicate_index.append(i+1)
    # for line in data2:
        # print(line)
    #    pass
    # print(duplicate_index)
    # for j in duplicate_index:
        # print(j)
    #    del data2[j]
    #df = pandas.DataFrame(data2,columns=['one', 'two', 'three'])
    #writer = pandas.ExcelWriter(path+filename, engine='xlsxwriter')
    #df.to_excel(writer, sheet_name='Added_Blank_Space', index=False)
    # writer.save()

    df2 = pandas.DataFrame(data, columns=['one', 'two', 'three'])
    writer2 = pandas.ExcelWriter(path+"temp_"+filename, engine='xlsxwriter')
    df2.to_excel(writer2, sheet_name='Added_Blank_Space', index=False)
    writer2.save()


def round(a, b):
    c = a-b
    if int(float(str(c).split(":")[-1].split(".")[-1])) > 5:
        return int(float(str(c).split(":")[-1].split(".")[0])) + 1
    else:
        return int(float(str(c).split(":")[-1].split(".")[0]))


def add_blank_rows_two(filename):

    df = pandas.read_excel(filename)
    df['difference'] = df.one.diff(1)
    print("DIFFERENCE COLUMN")
    print(df['difference'])

    list = df.values.tolist()
    list2 = []
    for i in list:
        if i[7] > datetime.timedelta(seconds=1.5):
            list2.append(["", "", "","","","",""])
            list2.append([i[0], i[1], i[2],i[3], i[4], i[5],i[6]])
        else:
            list2.append([i[0], i[1], i[2],i[3], i[4], i[5],i[6]])

    df = pandas.DataFrame(list2, columns=['one', 'two', 'three', "blank", "four", "five", "six"])
    writer = pandas.ExcelWriter(path+filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Added_Blank_Space', index=False)
    writer.save()


def combine_files(filename1, filename2):


    # TODO CODE HERE
    # ALGORITHM
    # FIRST COMBINE BOTH
    df1 = pandas.read_excel(path+filename1)
    df2 = pandas.read_excel(path+filename2)

    # DROPPING BLANK LINES FROM DF1 AND DF2

    df1.drop_duplicates(subset=['one'], keep=False, inplace=True)
    df2.drop_duplicates(subset=['one'], keep=False, inplace=True)

    print(df1)
    print(df2)
    s1 = pandas.merge(df1, df2, how='inner', on="one")
    #s1.replace({'NaT': None}).dropna(how="all")
    #s1['one'] = s1['one'].dt.round('S')
    s1.drop_duplicates(subset=['one'], keep=False, inplace=True)
    selected_date = s1[["one"]]
    s1.insert(3, 'blank', '')
    s1.insert(4, 'fourth', selected_date)
    #s1['six'] = s1['one']
    #s1['one'] = pandas.to_datetime(s1['one'])
    writer = pandas.ExcelWriter(
        "merge_sheet1_sheet2.xlsx", engine='xlsxwriter')
    s1.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    print("merged s1")
    print(s1)
    add_blank_rows_two("merge_sheet1_sheet2.xlsx")


    

def combine_files_two(filename1, filename2):


    # TODO CODE HERE
    # ALGORITHM
    # FIRST COMBINE BOTH
    df1 = pandas.read_excel(path+filename1)
    df2 = pandas.read_excel(path+filename2)

    print(df1)
    print(df2)
    s1 = pandas.DataFrame()
    
    s1['one'] = df1['one']
    s1['two'] = df1['two']
    s1['three'] = df1['three']
    
    s1['four'] = df2['one']
    s1['five'] = df2['two']
    s1['six'] = df2['three']
    
    
    s1.insert(3, 'blank', '')
    
    
    writer = pandas.ExcelWriter(
        path+"merge_sheet1_sheet2_two.xlsx", engine='xlsxwriter')
    s1.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    



def merge_converter(filename):
    df = pandas.read_excel(filename)
    df.insert(7, 'seven', '')
    # PERCENTAGE CALCULATIONS
    df['i'] = df['one']
    df['j'] = df['two']
    df['j'] = df['j'].fillna(0)
    df['k'] = df.j.diff()
    df['j'] = df['j'].replace(0,np.NaN)
    df['l'] = df['three']

    df['n'] = df['one']
    df['o'] = df['two']
    df['o'] = df['o'].fillna(0)
    df['p'] = df.o.diff()
    df['o'] = df['o'].replace(0,np.NaN)
    df['q'] = df['three']
    
    df.insert(12, 'twelve', '')

    df_one = pandas.Dataframe()

    


    
    writer = pandas.ExcelWriter(path+"merge_converter.xlsx", engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    print(df)


filenames = ["sheet1.xlsx", "sheet2.xlsx"]
for filename in filenames:
    if not os.path.exists(filename):
        print("Both Sheets names should be sheet1 and sheet2")
        input("Press Enter to exit the program")
        sys.exit()
    else:
        pass
for filename in filenames:
    print(filename)
    copyfile(filename, path+"temp_"+filename)
    try:
        pass
        # add_blank_rows(filename)
    except Exception as e:
        print(f"Add Blank Rows error of {filename}")
        print(str(e))
    try:
        main(filename)
    except Exception as e:
        print(f"Main Function error of {filename}")
        print(str(e))
try:
    #pass
    combine_files(filenames[0], filenames[1])
except Exception as e:
    print(f"Combine_files error")
    print(str(e))

try:
    # pass
    combine_files_two(filenames[0], filenames[1])
except Exception as e:
    print(f"Combine_files error")
    print(str(e))

if os.path.exists(path+"merge_sheet1_sheet2.xlsx"):
    try:
        #pass
        merge_converter(path+"merge_sheet1_sheet2.xlsx")
    except Exception as e:
        print(f"Merge Converter error")
        print(str(e))


if os.path.exists(path+"temp_"+filenames[0]):
    os.remove(path+"temp_"+filenames[0])

if os.path.exists(path+"temp_"+filenames[1]):
    os.remove(path+"temp_"+filenames[1])

if os.path.exists("merge_sheet1_sheet2.xlsx"):
    os.remove("merge_sheet1_sheet2.xlsx")
print("*****************************************")
print("*****************************************")

print("                       ")
print("                       ")

print("File converted successfully!!")
print("Press Enter to Continue!")
print("                       ")
print("                       ")

print("*****************************************")
print("*****************************************")
input()
sys.exit()
