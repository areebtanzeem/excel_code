import os,glob
import sys,re
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

path = join(dirname(realpath(__file__)), 'files/')
final_path = join(dirname(realpath(__file__)), 'final_files/')


#main_count = 0
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
    #global main_count
    #main_count = main_count + 1
    #print(f"Running Main Function for filename {filename} ")
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
    #print(f"Running Combine Function for filename {filename2} ")
    s2 = pandas.DataFrame(list2, columns=['one', 'two', 'three', "blank", "four", "five", "six"])
    
    writer = pandas.ExcelWriter(final_path+f"{f2}.xlsx", engine='xlsxwriter')
    s2.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    




def run_in_parallel_main():
    try:

        pool = Pool(processes=len(csv_filenames))
        pool.map(main, csv_filenames)
        pool.close()
        pool.join()
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))    


#MAIN FUNCTION ENDS HERE
print(os.path.dirname(os.path.realpath(__file__)))



#MERGING FILES START HERE



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

    #additional1 = pandas.DataFrame({'two_p_values': two_p_values})
    df = pandas.concat([df, df1], axis=1)
    #df['l'] = df1['l']
    #df['m'] = df1['m']
    #df['n'] = df1['n']
    #df['14'] = df1['14']
    #df['p'] = df1['p']
    #df['q'] = df1['q']
    #df['r'] = df1['r']

    df.insert(10,'10','')
    #print("THIS IS DF1")
    #print(df1)
    #df1.drop_duplicates(subset=['l'], inplace=True)
    #df1.reset_index(inplace=True,drop=True)
    #print(df1)

    #LMN STARTS HERE

    df["m"] = pandas.to_numeric(df["m"])
    df["n"] = pandas.to_numeric(df["n"])

    df['t'] = df.m.diff()
    df['temp_t'] = df.apply(lambda x : x['t'] if x['t'] > 0 else np.NaN , axis = 1)
    df['t'] = df['temp_t']
    del df['temp_t']
    df['t'] = df['t'].replace(np.NaN,0)
    df['u'] = df.apply(lambda x: x['n'] if x['t'] != 0 else np.NaN , axis = 1)
    

    df['prev_m'] = df['m'].shift()
    df['v'] = df.apply(lambda x : x['m']*100/x['prev_m'] if x['t'] != 0  and x['prev_m'] != 0 else np.NaN , axis = 1)
    
    #

    df['prev_n'] = df['n'].shift()
    df['w'] = df.apply(lambda x : x['n']*100/x['prev_n'] if x['t'] != 0 and x['prev_n'] != 0 else np.NaN , axis = 1)
    df['t'] = df['t'].replace(0,np.NaN)
    #

    df.insert(18,'18','')

    df['y'] = df.m.diff()
    df['temp_y'] = df.apply(lambda x : x['y'] if x['y'] < 0 else np.NaN , axis = 1)
    df['y'] = df['temp_y']
    del df['temp_y']
    df['y'] = df['y'].replace(np.NaN,0)
    df['z'] = df.apply(lambda x: x['n'] if x['y'] != 0 else np.NaN , axis = 1)

    df['aa'] = df.apply(lambda x : x['m']*100/x['prev_m'] if x['y'] != 0 and x['prev_m'] != 0 else np.NaN , axis = 1)
    df['ab'] = df.apply(lambda x : x['n']*100/x['prev_n'] if x['y'] != 0 and x['prev_n'] != 0 else np.NaN , axis = 1)

    df['y'] = df['y'].replace(0,np.NaN)
    df['y'] = df['y'].abs()
    del df['prev_m']
    del df['prev_n']
    df.insert(23,'23','')


    # LMN COMPLETES HERE

    # PQR STARTS HERE

    df["q"] = pandas.to_numeric(df["q"])
    df["r"] = pandas.to_numeric(df["r"])

    df['ad'] = df.q.diff()
    df['temp_ad'] = df.apply(lambda x : x['ad'] if x['ad'] > 0 else np.NaN , axis = 1)
    df['ad'] = df['temp_ad']
    del df['temp_ad']
    df['ad'] = df['ad'].replace(np.NaN,0)
    df['ae'] = df.apply(lambda x: x['r'] if x['ad'] != 0 else np.NaN , axis = 1)
    

    df['prev_q'] = df['q'].shift()
    df['af'] = df.apply(lambda x : x['q']*100/x['prev_q'] if x['ad'] != 0 and x['prev_q'] != 0 else np.NaN , axis = 1)
    
    #

    df['prev_r'] = df['r'].shift()
    df['ag'] = df.apply(lambda x : x['r']*100/x['prev_r'] if x['ad'] != 0 and x['prev_r'] != 0 else np.NaN , axis = 1)
    df['ad'] = df['ad'].replace(0,np.NaN)
    #

    df.insert(28,'28','')

    df['ai'] = df.q.diff()
    df['temp_ai'] = df.apply(lambda x : x['ai'] if x['ai'] < 0 else np.NaN , axis = 1)
    df['ai'] = df['temp_ai']
    del df['temp_ai']
    df['ai'] = df['ai'].replace(np.NaN,0)
    df['aj'] = df.apply(lambda x: x['r'] if x['ai'] != 0 else np.NaN , axis = 1)

    df['ak'] = df.apply(lambda x : x['q']*100/x['prev_q'] if x['ai'] != 0 and x['prev_q'] != 0 else np.NaN , axis = 1)
    df['al'] = df.apply(lambda x : x['r']*100/x['prev_r'] if x['ai'] != 0 and x['prev_r'] != 0 else np.NaN , axis = 1)

    df['ai'] = df['ai'].replace(0,np.NaN)
    df['ai'] = df['ai'].abs()
    del df['prev_q']
    del df['prev_r']
    df.insert(33,'33','')

    f_name = filename.split('.')[0]
    final_sheet_data.append([f_name, df['two'].count(), df['two_diff'].count(), np.NaN , df['t'].sum() ,  df['u'].sum() , df['v'].sum() , df['w'].sum() , np.NaN,  df['y'].sum() , df['z'].sum() , df['aa'].sum() , df['ab'].sum() , np.NaN,  df['ad'].sum() , df['ae'].sum() , df['af'].sum() , df['ag'].sum() , np.NaN,  df['ai'].sum() , df['aj'].sum() , df['ak'].sum() , df['al'].sum() ])


    temp_df = pandas.DataFrame()
    temp_df['t'] = df['t']
    temp_df['u'] = df['u']
    temp_df['v'] = df['v']
    temp_df['w'] = df['w']
    temp_df = temp_df.dropna(how='all')
    temp_df.reset_index(inplace = True)
    temp_df.loc[temp_df.shape[0]]= temp_df.sum(numeric_only=True, axis=0)
    #print("this is tempdf", temp_df)
    df['t'] = temp_df['t']
    df['u'] = temp_df['u']
    df['v'] = temp_df['v']
    df['w'] = temp_df['w']
    del temp_df

    temp_df = pandas.DataFrame()
    temp_df['y'] = df['y']
    temp_df['z'] = df['z']
    temp_df['aa'] = df['aa']
    temp_df['ab'] = df['ab']
    temp_df = temp_df.dropna(how='all')
    temp_df.reset_index(inplace = True)
    temp_df.loc[temp_df.shape[0]]= temp_df.sum(numeric_only=True, axis=0)
    #print("this is tempdf", temp_df)
    df['y'] = temp_df['y']
    df['z'] = temp_df['z']
    df['aa'] = temp_df['aa']
    df['ab'] = temp_df['ab']
    del temp_df

    temp_df = pandas.DataFrame()
    temp_df['ad'] = df['ad']
    temp_df['ae'] = df['ae']
    temp_df['af'] = df['af']
    temp_df['ag'] = df['ag']
    temp_df = temp_df.dropna(how='all')
    temp_df.reset_index(inplace = True)
    temp_df.loc[temp_df.shape[0]]= temp_df.sum(numeric_only=True, axis=0)
    #print("this is tempdf", temp_df)
    df['ad'] = temp_df['ad']
    df['ae'] = temp_df['ae']
    df['af'] = temp_df['af']
    df['ag'] = temp_df['ag']
    del temp_df

    temp_df = pandas.DataFrame()
    temp_df['ai'] = df['ai']
    temp_df['aj'] = df['aj']
    temp_df['ak'] = df['ak']
    temp_df['al'] = df['al']
    temp_df = temp_df.dropna(how='all')
    temp_df.reset_index(inplace = True)
    temp_df.loc[temp_df.shape[0]]= temp_df.sum(numeric_only=True, axis=0)
    #print("this is tempdf", temp_df)
    df['ai'] = temp_df['ai']
    df['aj'] = temp_df['aj']
    df['ak'] = temp_df['ak']
    df['al'] = temp_df['al']
    del temp_df
    #df.insert(38,'38','')
    #df = pandas.concat([df, df1], axis=1)

    #PQR ENDS HERE
    
    #APPENDING DATA FOR FINAL SHEET
    
    #print("MERGE CONVERTER FOR FILENAME", filename)

    # DELETING P Q R ROWS FROM DF1

    df1.drop(columns=['14', 'p','q','r'] , inplace=True)
    print(df1)

    writer2 = pandas.ExcelWriter('step2/main.xlsx', engine='xlsxwriter') 
    df1.to_excel(writer2, sheet_name='Sheet1', index=False)
    writer2.save()

    writer = pandas.ExcelWriter(final_path+f"merge_"+filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()



def run_in_parallel_merge_converter():
    try:
        pool = Pool(processes=len(final_files))
        pool.map(merge_converter,final_files)
        pool.close()
        pool.join()
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))

def final_merge(filenames,final_sheet_data):
    writer = pandas.ExcelWriter(final_path+"Final File.xlsx", engine='xlsxwriter')
    final_merge_count = 0
    for name in filenames:
        try:
            df = pandas.read_excel(final_path+name)
            df.to_excel(writer, sheet_name=name.split('.')[0], index=False)
        except Exception as e:
            print("!! EXCEPTION OCCURED !!")
            print(str(e))
            pass
        final_merge_count += 1
        percent_final_merge = (final_merge_count*100)/(len(filenames))
        print(f'Final Merge Function {percent_final_merge}% Completed!')


    df2 = pandas.DataFrame(final_sheet_data)
    df2.to_excel(writer, sheet_name="Final", index=False)
    writer.save()

#for item in final_sheet_data:
#    print(item)

if __name__ == '__main__':

    start_whole = time.process_time()

    if not os.path.exists("files"):
        os.mkdir("files")

    if not os.path.exists("final_files"):
        os.mkdir("final_files")

    if not os.path.exists("step2"):
        os.mkdir("step2")

    

    filenames = glob.glob('*.txt')
    #print('FIlenames', filenames)
    for file in filenames:
        try:
            os.rename(file,file.split('.')[0]+'.csv')
        except Exception as e:
            print("!! EXCEPTION OCCURED !!")
            print(str(e))
            pass


    ####    MAIN FUNCTION TAKES CSV_FILENAMES #######
    csv_filenames = glob.glob('*.csv')
    #print("CSV FILENAMES ", csv_filenames)
    start = time.process_time()
    #run_in_parallel_main()
    main_count = 0
    for i in csv_filenames:
        try:
            main(i)
        except Exception as e:
            print("!! EXCEPTION OCCURED !!")
            print(str(e))
            pass


        main_count += 1
        percent_main = (main_count*100)/(len(csv_filenames))
        print(f'Main Function completed {percent_main}%')
    
    print("TIME TAKEN TO EXECUTE main function",time.process_time() - start)

    
    ####       COMBINE FUNCTION TAKES XLSX FILENAMES  #####
    os.chdir('./files')
    xlsx_filenames = glob.glob('*.xlsx')
    
    try:
        xlsx_filenames.remove('main.xlsx')
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))
        pass
    
    os.chdir("../")
    #print('XLSX FILENAMES ', xlsx_filenames)
    
    start_combine = time.process_time()    
    #run_in_parallel_combine()
    combine_count = 0
    for j in xlsx_filenames:
        try:
            combine_files(j)
        except Exception as e:
            print("!! EXCEPTION OCCURED !!")
            print(str(e))
            pass


        combine_count += 1
        percent_combine = (combine_count*100)/(len(xlsx_filenames))
        print(f'Combine Function completed {percent_combine}%')
    
    print("TIME TAKEN TO EXECUTE Combine function",time.process_time() - start_combine)
    
    
    ###    MERGE CONVERTER FUNCTION TAKES FINAL FILES  #####
    os.chdir('./final_files')

    final_files = glob.glob('*.xlsx')

    os.chdir("../")
    try:
        final_files = sorted(final_files)
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))
        pass
    
    #print("Final Files",final_files)
    
    final_sheet_data = []


    
    start_merge_converter = time.process_time()
    #run_in_parallel_merge_converter()
    merge_count = 0
    for k in final_files:
        try:
            merge_converter(k)
        except Exception as e:
            print("!! EXCEPTION OCCURED !!")
            print(str(e))
            pass


        merge_count += 1
        percent_merge = (merge_count*100)/(len(final_files))
        print(f'Merge Function completed {percent_merge}%')
    
    
    print("TIME TAKEN TO EXECUTE merge_converter function",time.process_time() - start_merge_converter)

    #for item in final_sheet_data:
    #    print(item)
    
    #DELETING FILES IN FINAL_FILES LIST

    for name in final_files:
        if os.path.exists(final_path+name):
            os.remove(final_path+name)

    os.chdir('./final_files')

    final_merge_files = glob.glob('*.xlsx')
    try:
        final_merge_files = sorted(final_merge_files)
    except Exception as e:
        print("!! EXCEPTION OCCURED !!")
        print(str(e))
        pass


    os.chdir("../")

    #final_files.sort()
    #print("Final Merge Files",final_merge_files) 

    #### FINAL FUNCTION TAKES FINAL_MERGE FILE AS A LIST  #####

    final_merge(final_merge_files,final_sheet_data)
    for name in final_merge_files:
        if os.path.exists(final_path+name):
            os.remove(final_path+name)

    print("*****************************************")
    print("*****************************************")

    print("                       ")
    print("                       ")

    print("Total Time taken ",time.process_time() - start_whole,'s')

    print("File converted successfully!!")
    print("Press Enter to Continue!")
    print("                       ")
    print("                       ")

    print("*****************************************")
    print("*****************************************")
    input()
    sys.exit()
