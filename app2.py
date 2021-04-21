import os,glob
import sys
from os.path import join, dirname, realpath
import pandas
from openpyxl import load_workbook
import datetime
from shutil import copyfile
import numpy as np
import operator
import math
import sys, time, threading
from multiprocessing import Pool
import tqdm

if not os.path.exists("files"):
    os.mkdir("files")

if not os.path.exists("final_files"):
    os.mkdir("final_files")

path = join(dirname(realpath(__file__)), 'files/')
final_path = join(dirname(realpath(__file__)), 'final_files/')

filenames = glob.glob('*.txt')
for file in filenames:
    os.rename(file,file.split('.')[0]+'.csv')

main_count = 0
def main(filename):
    df = pandas.read_csv(filename, names=['one', 'two', 'three','four'],header=None)
    del df['one']
    df = df.rename(columns={"two": "one", "three": "two","four":"three"})
    df['one'] = pandas.to_datetime(df['one'])
    #print(df)
    df2 = df.groupby(['one'], as_index=False).agg(
        {'two': 'mean', 'three': 'sum'})

    # For B column average and for C column sun Ends Here

    #df.drop_duplicates(subset=['one'], keep=False, inplace=True)
    #print(df2.dtypes)
    df2['difference'] = df2.one.diff(1)
    list = df2.values.tolist()
    list2 = []
    for i in list:
        if i[3] > datetime.timedelta(seconds=1):
            list2.append(["", "", ""])
            list2.append([i[0], i[1], i[2]])
        else:
            list2.append([i[0], i[1], i[2]])
    global main_count
    main_count = main_count + 1
    print(f"Running Main Function for filename {filename} ")
    df3 = pandas.DataFrame(list2, columns=['one', 'two', 'three'])
    filename = filename.split(".")[0]+'.xlsx'
    writer = pandas.ExcelWriter(path+filename, engine='xlsxwriter')
    df3.to_excel(writer, sheet_name='Sheet', index=False)
    writer.save()
    #print(df3)

def combine_files(filename2):
    filename1 = 'main.xlsx'
    f1 = filename1.split('.')[0]
    f2 = filename2.split('.')[0]
    # TODO CODE HERE
    # ALGORITHM
    # FIRST COMBINE BOTH
    df1 = pandas.read_excel(path+filename1)
    df2 = pandas.read_excel(path+filename2)

    # DROPPING BLANK LINES FROM DF1 AND DF2

    df1.drop_duplicates(subset=['one'], keep=False, inplace=True)
    df2.drop_duplicates(subset=['one'], keep=False, inplace=True)

    
    s1 = pandas.merge(df1, df2, how='inner', on="one")
    
    selected_date = s1[["one"]]
    s1.insert(3, 'blank', '')
    s1.insert(4, 'fourth', selected_date)

    #BLANK ROWS STARTS FROM HERE

    s1['difference'] = s1.one.diff(1)
    list = s1.values.tolist()
    list2 = []
    for i in list:
        if i[7] > datetime.timedelta(seconds=1.5):
            list2.append(["", "", "","","","",""])
            list2.append([i[0], i[1], i[2],i[3], i[4], i[5],i[6]])
        else:
            list2.append([i[0], i[1], i[2],i[3], i[4], i[5],i[6]])
    print(f"Running Main Function for filename {filename2} ")
    s2 = pandas.DataFrame(list2, columns=['one', 'two', 'three', "blank", "four", "five", "six"])
    
    writer = pandas.ExcelWriter(final_path+f"{f2}.xlsx", engine='xlsxwriter')
    s2.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    


start = time.process_time()
csv_filenames = glob.glob('*.csv')
def run_in_parallel_main():
    try:

        pool = Pool(processes=len(csv_filenames))
        pool.map(main, csv_filenames)
        pool.close()
        pool.join()
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))    


#run_in_parallel_main()

print("TIME TAKEN TO EXECUTE main function",time.process_time() - start)

#MAIN FUNCTION ENDS HERE
print(os.path.dirname(os.path.realpath(__file__)))

os.chdir('./files')
xlsx_filenames = glob.glob('*.xlsx')

xlsx_filenames.remove('main.xlsx')
os.chdir("../")

#MERGING FILES START HERE


start_combine = time.process_time()
#for i in xlsx_filenames:
def run_in_parallel_combine():
    try:
        pool = Pool(processes=len(xlsx_filenames))
        pool.map(combine_files,xlsx_filenames)
        pool.close()
        pool.join()
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))

#run_in_parallel_combine()
print("TIME TAKEN TO EXECUTE Combine function",time.process_time() - start_combine)

os.chdir('./final_files')

final_files = glob.glob('*.xlsx')

os.chdir("../")
print(final_files)

def merge_converter(filename):
    df = pandas.read_excel(final_path+filename)
    df['two_p_diff'] = df.two.diff()
    df['five_p_diff'] = df.five.diff()
    df.insert(7,'seven','')

    #ONLY REMAIN SAME SIGNS

    df['two_diff'] = df.apply(lambda x: x['two_p_diff'] if x['two_p_diff']*x['five_p_diff'] > 0 else np.NaN, axis=1)
    df['five_diff'] = df.apply(lambda x: x['five_p_diff'] if x['two_p_diff']*x['five_p_diff'] > 0 else np.NaN, axis=1)

    del df['five_p_diff']
    del df['two_p_diff']

    index_list = df[~df.two_diff.isnull()].index.tolist()
    two_data = []

    for j in index_list:
        two_data.append([np.NaN,np.NaN,np.NaN,np.NaN,np.NaN,np.NaN,np.NaN])
        two_data.append( [df._get_value(j-1, 'one') , df._get_value(j-1, 'two') , df._get_value(j-1, 'three') , df._get_value(j-1, 'blank') , df._get_value(j-1, 'four') , df._get_value(j-1, 'five') , df._get_value(j-1, 'six')  ] )
        two_data.append( [df._get_value(j, 'one') , df._get_value(j, 'two') , df._get_value(j, 'three') , df._get_value(j, 'blank') , df._get_value(j, 'four') , df._get_value(j, 'five') , df._get_value(j, 'six')  ] )

    #print(two_data)

    columns = ['l', 'm', 'n','14','p','q','r']
    df1 = pandas.DataFrame(two_data,columns=columns)
    #print("THIS IS Df!")
    #print(df1)

    df['l'] = df1['l']
    df['m'] = df1['m']
    df['n'] = df1['n']
    df['14'] = df1['14']
    df['p'] = df1['p']
    df['q'] = df1['q']
    df['r'] = df1['r']

    df.insert(10,'10','')
    df["m"] = pandas.to_numeric(df["m"])
    df["n"] = pandas.to_numeric(df["n"])

    df['t'] = df.m.diff()
    df['t'] = df['t'].replace(np.NaN,0)
    df['u'] = df.apply(lambda x: x['n'] if x['t'] != 0 else np.NaN , axis = 1)
    

    df['prev_m'] = df['m'].shift()
    df['v'] = df.apply(lambda x : x['m']*100/x['prev_m'] if x['t'] != 0 else np.NaN , axis = 1)
    df['t'] = df['t'].replace(0,np.NaN)
    del df['prev_m']

    df.insert(18,'18','')

    

    writer = pandas.ExcelWriter(final_path+f"merge_"+filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()

merge_converter('A.xlsx')