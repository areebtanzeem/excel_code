import os
import sys
from os.path import join, dirname, realpath
import pandas
from openpyxl import load_workbook
import datetime
from shutil import copyfile
import numpy as np

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
    df = pandas.read_excel(filename, names=['one', 'two', 'three'])
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
    data = []
    data2 = []
    gap_rows = []
    duplicate_index = []
    workbook = load_workbook(filename=filename)
    sheet = workbook['Sheet1']
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)

    for i in range(0, len(data)-1):
        a = data[i+1][0]
        b = data[i][0]
        diff = round(a, b)
        print(data[i][0])
        print(data[i+1][0])
        print(diff, "   time delta   ", datetime.timedelta(seconds=1.5))

        data2.append(list(data[i]))
        if diff > 1:
            data2.append(list(("", "", "")))
        elif diff == 0:
            duplicate_index.append(i)
            duplicate_index.append(i+1)
    for line in data2:
        # print(line)
        pass
    print(duplicate_index)
    for j in duplicate_index:
        del data2[j]
    df = pandas.DataFrame(
        data2, columns=['one', 'two', 'three', "blank", "four", "five", "six"])
    writer = pandas.ExcelWriter(path+filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Added_Blank_Space', index=False)
    writer.save()


def combine_files(filename1, filename2):

    #workbook1 = load_workbook(filename=filename1)
    #sheet1 = workbook1['Added_Blank_Space']

    #workbook2 = load_workbook(filename=filename2)
    #sheet2 = workbook2['Added_Blank_Space']

    #data1 = []
    #data2 = []
    #final_data = []
    #data1_index = []
    #data2_index = []
    #print("data1 length",len(data1))
    #print("data2 length",len(data2))
    # for row1 in sheet1.iter_rows(values_only=True):
    #   data1.append(list(row1))

   # for row2 in sheet2.iter_rows(values_only=True):
   #     data2.append(list(row2))

    #    for i in range(len(data1)-1):
    #        for j in range(len(data2)-1):
    #            if data1[i][0] in data2[j]:
    #                pass
    #            else:
    #                data1_index.append(i)
    #                data2_index.append(j)

    #    print(data2_index)
    #    print(data1_index)

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
    print(s1)
    add_blank_rows_two("merge_sheet1_sheet2.xlsx")


def merge_converter(filename):
    df = pandas.read_excel(filename)
    #df.drop_duplicates(subset=['one'], keep=False, inplace=True)
    #df.loc['Total'] = pandas.Series(df.sum())
    df.insert(7, 'seven', '')
    # PERCENTAGE CALCULATIONS
    df['two_percentage'] = df['two'].apply(lambda a: (a/df['two'].sum())*100)
    df['five_percentage'] = df['five'].apply(
        lambda a: (a/df['five'].sum())*100)
    df['three_percentage'] = df['three'].apply(
        lambda a: (a/df['three'].sum())*100)
    
    df['six_percentage'] = df['six'].apply(lambda a: (a/df['six'].sum())*100)

    # PERCENTAGE CALCULATED

    #df['sum_two_three_percent'] = df.apply(lambda row: row.two_percentage + row.three_percentage, axis=1)
    #df['sum_five_six_percent'] = df.apply(lambda row: row.five_percentage + row.six_percentage, axis=1)
    # Y or AB ka difference lena hai means b and f ka percentage means two and five ka diff
    # DIFFERENCE

    df['two_p_diff'] = df.two_percentage.diff()
    df['five_p_diff'] = df.five_percentage.diff()
    df['three_p_diff'] = df.three_percentage.diff()
    df['six_p_diff'] = df.six_percentage.diff()

    # DIFFERENCE ENDS

    # SAME SAME

    df['two_p_same'] = df.apply(lambda x: x['two_p_diff'] if x['two_p_diff']*x['five_p_diff'] > 0 else np.NaN, axis=1)
    df['five_p_same'] = df.apply(lambda x: x['five_p_diff'] if x['five_p_diff']*x['two_p_diff'] > 0 else np.NaN, axis=1)
    df['three_p_same'] = df.apply(lambda x: x['three_p_diff'] if x['three_p_diff']*x['six_p_diff'] > 0 else np.NaN, axis=1)
    df['six_p_same'] = df.apply(lambda x: x['six_p_diff'] if x['three_p_diff']*x['six_p_diff'] > 0 else np.NaN, axis=1)

    # SaME SAME

    #Changing p_diff to p_same and deleting  same

    df['two_p_diff'] = df['two_p_same']
    df['five_p_diff'] = df['five_p_same']
    df['three_p_diff'] = df['three_p_same']
    df['six_p_diff'] = df['six_p_same']

    del df['two_p_same']
    del df['five_p_same']
    del df['three_p_same']
    del df['six_p_same']
    
    ## CHECK IF ALL FOUR COLUMNS HAVE DATA IN IT OTHERWISE ENTER NULL DATA

    df.loc[df.three_p_diff.isnull(), ['two_p_diff','five_p_diff']] = np.NaN
    df.loc[df.two_p_diff.isnull(), ['three_p_diff','six_p_diff']] = np.NaN


    #GET VALUE OF ABOVE ROW FOR TWO PERCENTAGE

    index_of_not_null_two_p_diff = df[~df.two_p_diff.isnull()].index.tolist()
    print(index_of_not_null_two_p_diff)

    sum_above_two_diff_pos = 0
    sum_above_three_diff_pos = 0
    sum_above_five_diff_pos = 0
    sum_above_six_diff_pos = 0

    sum_same_two_diff_pos = 0
    sum_same_three_diff_pos = 0
    sum_same_five_diff_pos = 0
    sum_same_six_diff_pos = 0

    sum_above_two_diff_neg = 0
    sum_above_three_diff_neg = 0
    sum_above_five_diff_neg = 0
    sum_above_six_diff_neg = 0

    sum_same_two_diff_neg = 0
    sum_same_three_diff_neg = 0
    sum_same_five_diff_neg = 0
    sum_same_six_diff_neg = 0


    for i in index_of_not_null_two_p_diff:
        print(f"VALUE OF INDEX TWO DIFFERENCE FOR INDEX {i-1}:  ",df._get_value(i-1, 'two_percentage'))
        print(f"VALUE OF INDEX TWO DIFFERENCE FOR INDEX {i}:  ",df._get_value(i, 'two_percentage'))
        
        if df._get_value(i, 'three_p_diff') > 0:

            sum_above_two_diff_pos += df._get_value(i-1, 'two_percentage')
            sum_above_three_diff_pos += df._get_value(i-1, 'three_percentage')
            sum_above_five_diff_pos += df._get_value(i-1, 'five_percentage')
            sum_above_six_diff_pos += df._get_value(i-1, 'six_percentage')


            sum_same_two_diff_pos += df._get_value(i, 'two_percentage')
            sum_same_three_diff_pos += df._get_value(i, 'three_percentage')
            sum_same_five_diff_pos += df._get_value(i, 'five_percentage')
            sum_same_six_diff_pos += df._get_value(i, 'six_percentage')
        elif df._get_value(i, 'three_p_diff') < 0:

            sum_above_two_diff_neg += df._get_value(i-1, 'two_percentage')
            sum_above_three_diff_neg += df._get_value(i-1, 'three_percentage')
            sum_above_five_diff_neg += df._get_value(i-1, 'five_percentage')
            sum_above_six_diff_neg += df._get_value(i-1, 'six_percentage')


            sum_same_two_diff_neg += df._get_value(i, 'two_percentage')
            sum_same_three_diff_neg += df._get_value(i, 'three_percentage')
            sum_same_five_diff_neg += df._get_value(i, 'five_percentage')
            sum_same_six_diff_neg += df._get_value(i, 'six_percentage')


    #POSITIVE

    two_above_same_diff_pos = sum_same_two_diff_pos - sum_above_two_diff_pos
    three_above_same_diff_pos = sum_same_three_diff_pos - sum_above_three_diff_pos
    five_above_same_diff_pos = sum_same_five_diff_pos - sum_above_five_diff_pos
    six_above_same_diff_pos = sum_same_six_diff_pos - sum_above_six_diff_pos

    total_rows = df.shape[0] - 3
    blank_list = [np.NaN]*(total_rows)

    df['two_dif_sum_pos'] = [sum_above_two_diff_pos,sum_same_two_diff_pos,two_above_same_diff_pos]+blank_list
    df['three_dif_sum_pos'] = [sum_above_three_diff_pos,sum_same_three_diff_pos,three_above_same_diff_pos]+blank_list
    df['five_dif_sum_pos'] = [sum_above_five_diff_pos,sum_same_five_diff_pos,five_above_same_diff_pos]+blank_list
    df['six_dif_sum_pos'] = [sum_above_six_diff_pos,sum_same_six_diff_pos,six_above_same_diff_pos]+blank_list

    #NEGATIVE

    two_above_same_diff_neg = sum_same_two_diff_neg - sum_above_two_diff_neg
    three_above_same_diff_neg = sum_same_three_diff_neg - sum_above_three_diff_neg
    five_above_same_diff_neg = sum_same_five_diff_neg - sum_above_five_diff_neg
    six_above_same_diff_neg = sum_same_six_diff_neg - sum_above_six_diff_neg

    #total_rows = df.shape[0] - 3
    #blank_list = [np.NaN]*(total_rows)

    df['two_dif_sum_neg'] = [sum_above_two_diff_neg,sum_same_two_diff_neg,two_above_same_diff_neg]+blank_list
    df['three_dif_sum_neg'] = [sum_above_three_diff_neg,sum_same_three_diff_neg,three_above_same_diff_neg]+blank_list
    df['five_dif_sum_neg'] = [sum_above_five_diff_neg,sum_same_five_diff_neg,five_above_same_diff_neg]+blank_list
    df['six_dif_sum_neg'] = [sum_above_six_diff_neg,sum_same_six_diff_neg,six_above_same_diff_neg]+blank_list


    #SHOW ROWS TWO THREE FIVE SIX ABOVE AND SAME

    two_p_values = []
    three_p_values = []
    five_p_values = []
    six_p_values = []

    for j in index_of_not_null_two_p_diff:
        two_p_values.append('')
        two_p_values.append(df._get_value(j-1, 'two_percentage'))
        two_p_values.append(df._get_value(j, 'two_percentage'))

        three_p_values.append('')
        three_p_values.append(df._get_value(j-1, 'three_percentage'))
        three_p_values.append(df._get_value(j, 'three_percentage'))

        five_p_values.append('')
        five_p_values.append(df._get_value(j-1, 'five_percentage'))
        five_p_values.append(df._get_value(j, 'five_percentage'))

        six_p_values.append('')
        six_p_values.append(df._get_value(j-1, 'six_percentage'))
        six_p_values.append(df._get_value(j, 'six_percentage'))
    
    
    total_rows_another = df.shape[0]
    another_blank_list = [np.NaN]
    df['two_p_values'] = two_p_values + another_blank_list*(total_rows_another - len(two_p_values))
    df['three_p_values'] = three_p_values + another_blank_list*(total_rows_another - len(three_p_values))
    df['five_p_values'] = five_p_values + another_blank_list*(total_rows_another - len(five_p_values))
    df['six_p_values'] = six_p_values + another_blank_list*(total_rows_another - len(six_p_values))



    #df.loc[df['two_p_diff'] != "", 'above_two_p'] = df["two_percentage"].shift(-1)
    #df['above_two_p_test'] = df.loc[df['two_p_diff'] != np.NaN, 'A'.shift(-1)]

    #df['above_two_p'] = df.apply(lambda x: df['two_percentage'].shift(1) if ~df['two_p_diff'].isin(['']) else np.NaN,axis=1)



    # CONVERTING SAME SAME TO POSITIVE

    #df['sum_two_three_difference_positive'] = df['sum_two_three_difference_a'].apply(lambda x : abs(x) if x != "" else np.NaN )
    #df['sum_five_six_difference_positive'] = df['sum_five_six_difference_a'].apply(lambda x : abs(x) if x != "" else np.NaN )

    #df['sum_two_three_difference_a'] = df['sum_two_three_difference_positive']
    #df['sum_five_six_difference_a'] = df['sum_five_six_difference_positive']

    #del df['sum_two_three_difference_positive']
    #del df['sum_five_six_difference_positive']
    #print(df['sum_two_three_difference_a'])
    #print(df['sum_five_six_difference_a'])
    try:
        pass

        #sum_two_percentage_difference_a = df['sum_two_three_difference_a'].sum()
        #sum_five_percentage_difference_a = df['sum_five_six_difference_a'].sum()

        #percent_sum_two_percentage_difference_a = (sum_two_percentage_difference_a/(sum_two_percentage_difference_a+sum_five_percentage_difference_a))*100
        #percent_sum_five_percentage_difference_a = (sum_five_percentage_difference_a/(sum_two_percentage_difference_a+sum_five_percentage_difference_a))*100

        #s1 = pandas.Series(data = {0:sum_two_percentage_difference_a,2:percent_sum_two_percentage_difference_a})
        #s2 = pandas.Series(data = {0:sum_five_percentage_difference_a,2:percent_sum_five_percentage_difference_a})
        #df['pec_two'] = s1
        #df['pec_five'] = s2
        #print("##################################")
        ##print("##################################")
        #print("##################################")
        #print("##################################")
        # print(s1)
        # print(s2)
        # print([sum_five_percentage_difference_a,percent_sum_five_percentage_difference_a])
        #print("##################################")
        #print("##################################")
        #print("##################################")
        #print("##################################")

    except Exception as e:
        print(str(e))

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
    # pass
    combine_files(filenames[0], filenames[1])
except Exception as e:
    print(f"Combine_files error")
    print(str(e))

if os.path.exists(path+"merge_sheet1_sheet2.xlsx"):
    try:
        # pass
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