import pandas as pd
import shutil
import datetime
import os
import openpyxl

def runFilterLinux(input_path,output_path):
    
    df1 = pd.read_excel(input_path,'Server logon access')
    df2 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\company_list.xlsx')
    df3 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\department_list.xlsx')
    df4 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\employeetype_list.xlsx')
    df5 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\manager_list.xlsx')
    
    total_index = df1.index  
    total_num_entries = len(total_index)

    print("Applying filters for Employee Type")
    filtered_data = df1.loc[~(df1['EmployeeType'].isin(df4['EmployeeType']))]

    print("Applying filters for Company")
    filtered_data2 = filtered_data.loc[~(filtered_data['Company'].isin(df2['Company']))]

    print("Applying filters for Manager")
    filtered_data3 = filtered_data2.loc[(~(filtered_data2['L4manager'].isin(df5['Name']))) & (~(filtered_data2['L3manager'].isin(df5['Name'])))]
    filtered_data4 = filtered_data3.loc[(~(filtered_data3['L2manager'].isin(df5['Name']))) & (~(filtered_data3['L1manager'].isin(df5['Name'])))]
    filtered_data5 = filtered_data4.loc[~(filtered_data4['L0manager'].isin(df5['Name']))]

    print(f'Total number of entries without filtering {total_num_entries}')

    index = filtered_data5.index
    num_entries = len(index)
    print(f"Number of entries after filtering: {num_entries}")
    file_name1 = output_path + '\\IS Developers with Direct Access '+datetime.date.today().strftime('%m%d%Y')+'.xlsx'
    
    writer = pd.ExcelWriter(file_name1, engine='openpyxl',mode = 'a')
    filtered_data5.to_excel(writer, sheet_name = 'Direct Access', index=False)
    writer.save()
