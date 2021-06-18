import pandas as pd
import shutil
import datetime
import os
import openpyxl
from log import log
import datetime

today = datetime.date.today()
today = today.strftime('%m%d%Y')

def run_filter(input_path,output_path):
    
    df1 = pd.read_excel(input_path,'Users')
    df2 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\company_list.xlsx')
    df3 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\department_list.xlsx')
    df4 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\employeetype_list.xlsx')
    df5 = pd.read_excel('\\\\lausers\\trsusers01\\WeissmJ\\drivers\\Output\\SOX 2021\\IS Reference Docs\\SOX Automation\\manager_list.xlsx')
    
    total_index = df1.index  
    total_num_entries = len(total_index)

    print("Applying filters for Employee Type")
    filtered_data = df1.loc[~(df1['EmployeeType'].isin(df4['EmployeeType']))]

    print("Applying filters for Company")
    filtered_data1 = filtered_data.loc[~(filtered_data['Company'].isin(df2['Company']))]

    print("Applying filters for Department")
    filtered_data2 = filtered_data1.loc[~(filtered_data1['Department'].isin(df3['Department']))]

    print("Applying filters for Manager")
    filtered_data3 = filtered_data2.loc[(~(filtered_data2['L4manager'].isin(df5['Name']))) & (~(filtered_data2['L3manager'].isin(df5['Name'])))]
    filtered_data4 = filtered_data3.loc[(~(filtered_data3['L2manager'].isin(df5['Name']))) & (~(filtered_data3['L1manager'].isin(df5['Name'])))]
    filtered_data5 = filtered_data4.loc[~(filtered_data4['L0manager'].isin(df5['Name']))]

    print(f'Total number of entries without filtering {total_num_entries}')

    index = filtered_data5.index
    num_entries = len(index)
    print(f"Number of entries after filtering: {num_entries}")
    log('Direct Access entries found: '+str(num_entries))
    file_name1 = output_path + '\\IS Developers with Direct Access '+datetime.date.today().strftime('%m%d%Y')+'.xlsx'
    
    writer = pd.ExcelWriter(file_name1, engine='openpyxl',mode = 'a')
    filtered_data5.to_excel(writer, sheet_name = 'Direct Access', index=False)
    writer.save()

    # appending direct access found to consolidated files
    print('Generating the consolidated direct access file')
    consolidated_path = 'C:\\Python v2_sarthak\\Input Folder\\Concat Direct Access\\Direct Access Consolidated.xlsx'
    if os.path.isfile(consolidated_path):
        consolidated_file = pd.read_excel(consolidated_path,'Direct Access Consolidated')
        os.remove(consolidated_path)
        consolidated_direct_access_list = [filtered_data5,consolidated_file]
        result = pd.concat(consolidated_direct_access_list)
        result.to_excel(consolidated_path,sheet_name='Direct Access Consolidated',index=False)
    else:
        filtered_data5.to_excel(consolidated_path, sheet_name='Direct Access Consolidated', index=False)


    # appending server to direct access
    serv_to_direct_path = 'C:\\Python v2_sarthak\\Input Folder\\Concat Direct Access\\Server to Direct Access.xlsx'
    if os.path.isfile(consolidated_path):
        df1 = pd.read_excel(consolidated_path,'Direct Access Consolidated')
        output_df = pd.DataFrame(columns = ['Server Name','Users','User Count','Date'])
        length = len(df1.index)

        server_list = []
        for i in range(length):
            server = df1.at[i,'Server Name']
            if server not in server_list:
                server_list.append(server)
                user_list = []
                for j in range(i,length):
                    server_name = df1.at[j,'Server Name']
                    if server == server_name:
                        user = df1.at[j,'Name']
                        if user not in user_list:
                            user_list.append(user)
                separator = ';'
                users = separator.join(user_list)
                output_df = output_df.append({'Server Name': server,'Users': users,'User Count':len(user_list),'Date':today},ignore_index=True)
        if os.path.isfile(serv_to_direct_path):
            os.remove(serv_to_direct_path)
            output_df.to_excel(serv_to_direct_path,sheet_name='Serv to Direct Access',index=False)
        else:
            output_df.to_excel(serv_to_direct_path,sheet_name='Serv to Direct Access',index=False)


    # filtering for group members

    df1 = pd.read_excel(input_path,'Group members')

    total_index = df1.index  
    total_num_entries = len(total_index)

    print("Applying filters for Employee Type")
    filtered_data = df1.loc[~(df1['EmployeeType'].isin(df4['EmployeeType']))]

    print("Applying filters for Company")
    filtered_data1 = filtered_data.loc[~(filtered_data['Company'].isin(df2['Company']))]

    print("Applying filters for Department")
    filtered_data2 = filtered_data1.loc[~(filtered_data1['Department'].isin(df3['Department']))]

    print("Applying filters for Manager")
    filtered_data3 = filtered_data2.loc[(~(filtered_data2['L4manager'].isin(df5['Name']))) & (~(filtered_data2['L3manager'].isin(df5['Name'])))]
    filtered_data4 = filtered_data3.loc[(~(filtered_data3['L2manager'].isin(df5['Name']))) & (~(filtered_data3['L1manager'].isin(df5['Name'])))]
    filtered_data5 = filtered_data4.loc[~(filtered_data4['L0manager'].isin(df5['Name']))]

    print(f'Total number of entries without filtering {total_num_entries}')

    index = filtered_data5.index
    num_entries = len(index)
    print(f"Number of entries after filtering: {num_entries}")
    log('Developers with access through AD Group: '+str(num_entries))
    file_name2 = output_path + '\\IS Developers with Access thru AD Group '+datetime.date.today().strftime('%m%d%Y')+'.xlsx'
    
    writer = pd.ExcelWriter(file_name2, engine='openpyxl',mode = 'a')
    filtered_data5.to_excel(writer, sheet_name = 'IS Dev with Access thru AD Grp', index=False)
    writer.save()

    # creating the 3rd Complete file
    print('Applying vlookup')
    df1 = pd.read_excel(output_path+'\\IS Developers with Access thru AD Group '+datetime.date.today().strftime('%m%d%Y')+'.xlsx','IS Dev with Access thru AD Grp')
    df2 = pd.read_excel(input_path,'Groups')
   
    user_access_through_ad = len(df1.index)
    group_perm = len(df2.index)

    master_list = []

    for i in range(user_access_through_ad):
        SamAccountName = df1.at[i,'SamAccountName']
        server_list = []
        perm_list = []
        for j in range(group_perm):
            SamAccountName_perm = df2.at[j,'Group/User']
            SamAccountName_perm = SamAccountName_perm.split('\\')[1]
            if SamAccountName==SamAccountName_perm:
                server_name = df2.at[j,'Server Name']
                permission = df2.at[j,'Permissions']
                if server_name not in server_list:
                    server_list.append(server_name)
                if permission not in perm_list:
                    perm_list.append(permission)
        separator = ', '
        if len(perm_list)>1:
            perm = 'Different permissions on different paths, see Group Permissions'
        else:
            perm = perm_list[0]
        df1.at[i,'Permissions'] = perm
        server = separator.join(server_list)
        df1.at[i,'Server'] = server


    cols = ['Server']  + ['SamAccountName'] + ['Permissions'] + [col for col in df1 if col not in ['Server','SamAccountName','Permissions']]
    if 'Server' in df1:
        df1 = df1[cols]
    df1 = (df1.style.applymap(lambda v: 'background-color: %s' % 'yellow' if v=='Different permissions on different paths, see Group Permissions' else ''))
    file_name3 = output_path + '\\IS Dev with Access thru AD Group '+datetime.date.today().strftime('%m%d%Y')+' - Complete.xlsx'
    df1.to_excel(file_name3,sheet_name='User Access through AD Groups',index=False)

    df3 = pd.DataFrame(columns = ['User/Group','Permissions'])
    df4 = pd.DataFrame(columns = ['User/Group','Permissions'])

    SamAccountName_list = []
    for j in range(group_perm):
        SamAccountName_perm = df2.at[j,'Group/User']
        SamAccountName_perm = SamAccountName_perm.split('\\')[1]
        if SamAccountName_perm not in SamAccountName_list:
            SamAccountName_list.append(SamAccountName_perm)
            permission_list = []
            for i in range(j,group_perm):
                SamAccountName = df2.at[i,'Group/User']
                SamAccountName = SamAccountName.split('\\')[1]
                if SamAccountName==SamAccountName_perm:
                    permission = df2.at[i,'Permissions']
                    if permission not in permission_list:
                        permission_list.append(permission)
            master_list.append([SamAccountName_perm,permission_list])

    for k in range(len(master_list)):
        if len(master_list[k][1])>1:
            df3 = df3.append({'User/Group':master_list[k][0],'Permissions':'Different permissions on different paths, see Group Permissions'},ignore_index=True)
            for n in range(len(master_list[k][1])):
                df4 = df4.append({'User/Group':master_list[k][0],'Permissions':master_list[k][1][n]},ignore_index=True)
        elif len(master_list[k][1])==1:
            df3 = df3.append({'User/Group':master_list[k][0],'Permissions':master_list[k][1][0]},ignore_index=True)
        else:
            df3 = df3.append({'User/Group':master_list[k][0],'Permissions':''},ignore_index=True)

    df3 = (df3.style.applymap(lambda v: 'background-color: %s' % 'yellow' if v=='Different permissions on different paths, see Group Permissions' else ''))
    df4 = (df4.style.applymap(lambda v: 'background-color: %s' % 'yellow'))
    print('Saving Files')
    writer = pd.ExcelWriter(file_name3, engine='openpyxl',mode = 'a')
    df2.to_excel(writer, sheet_name = 'Group Permissions', index=False)
    df3.to_excel(writer, sheet_name = 'AD Group Permissions -no dups', index=False)
    df4.to_excel(writer, sheet_name = 'AD groups with multiple perms', index=False)
    writer.save()
