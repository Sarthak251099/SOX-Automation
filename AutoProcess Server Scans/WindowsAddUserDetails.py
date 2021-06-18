# Reads in <n>_group-user.csv joins with active directory of people and groups/group members
# User IDs are resolved with manger details using active directory via GetUserGroup module
# Excel files produced using active directory date built up in LDAPcache.db

import pandas as pd
import sqlite3
import os
import Metrics
import GetUserGroup
from WriteExcel import Excel
from datetime import datetime


def addDetails(dataPath, metrics: Metrics.TrackWindows, env):
    print('Processing', dataPath)
    resultsPath = dataPath + '/results/'  # directory for output file results
    ad = GetUserGroup.Populate(env)

    # ############### Helper functions ###############

    def removeDomain(str1):
        # Strips off US domain name from user/group otherwise return original input
        if str1.startswith('US\\'):
            return str1.split('\\')[1]
        else:
            return str1

    # ############### Main ##########################

    # generate list of files to be processed
    filelist = [f for f in os.listdir(resultsPath) if os.path.isfile(os.path.join(resultsPath, f))]
    fileNumbers = []

    for file in filelist:
        if file.endswith('_group-user.csv'):
            fileNumbers.append(file.split('_')[0])

    # process each file - each will be smaller than the 1 million excel row limit from work done in WindowsACLscan.py
    for fileNumber in fileNumbers:
        aclInputFile = resultsPath + fileNumber + '_group-user.csv'
        finalOutputFile = resultsPath + 'group-user with AD info ' + fileNumber + '.xlsx'
        print('Input file:', aclInputFile)
        print('Output file:', finalOutputFile)

        # load file into pandas dataframe
        aclDF = pd.read_csv(aclInputFile)
        aclDF['SamAccountName'] = aclDF.apply(lambda x: removeDomain(x['Group/User']), axis=1)
        print('Length of ACL input file', len(aclDF))

        # Ensure LDAPcache db has a record for each line item
        rowcount = 0
        notFound = []
        for groupUser in aclDF['SamAccountName']:
            rowcount += 1
            # For performance see if they already exist in the user or group cache to avoid an unnecessary ldap call
            if not (ad.userInDB(groupUser) or ad.groupInDB(groupUser)):
                # Does not exist so try going back to ldap
                if not (ad.populateUserFromAD(groupUser) or ad.populateGroupFromAD(groupUser)):
                    if groupUser not in notFound:
                        notFound.append(groupUser)
            if rowcount % 10000 == 0:
                print(fileNumber, rowcount, datetime.now().strftime("%H:%M:%S"))
        print('Account records not found in Active Directory -', len(notFound), ':', notFound)

        # Build excel file
        workbook = Excel(finalOutputFile)

        # connect to db
        conn1 = sqlite3.connect(env + '_LDAPcache.db')

        # Users sheet
        userDF = pd.read_sql_query('select u.UserID as SamAccountName, "User" as Type, "" as Description, "" as Notes, u.CostCentre, u.Name, u.EmployeeType, u.Company, u.Department, m.L0manager, m.L1manager, m.L2manager, m.L3manager, m.L4manager '
                                   'from users u left join userToManagers m on u.DN=m.DN '
                                   'union '
                                   'select Name as SamAccountName, "Group" as Type, Description, Notes, "" as CostCentre, "" as Name, "" as EmployeeType, "" as Company, "" as Department, "" as L0manager, "" as L1manager, '
                                   '"" as L2manager, "" as L3manager, "" as L4manager from groups', conn1)

        excelDF = pd.merge(aclDF, userDF, how='left', on='SamAccountName')
        excelDF.drop(columns=['SamAccountName'], inplace=True)
        print('Length of Users output sheet', len(excelDF.query('(Type=="User")')))
        workbook.addSheet('Users', excelDF.query('(Type=="User")')[['Server Name', 'Folder Name', 'Is top level', 'Top Level', 'ACLs match top level', 'Group/User', 'Permissions', 'Applies To', 'Is Inherited', 'Type', 'CostCentre', 'Name',
                                                                    'EmployeeType','Company', 'Department', 'L0manager', 'L1manager', 'L2manager', 'L3manager', 'L4manager']])

        # Groups sheet
        print('Length of Groups output sheet', len(excelDF.query('(Type!="User")')))
        workbook.addSheet('Groups', excelDF.query('(Type=="Group")')[['Server Name', 'Folder Name', 'Top Level', 'Is top level', 'ACLs match top level', 'Group/User', 'Permissions', 'Applies To', 'Is Inherited', 'Type', 'Description', 'Notes']])

        # Store metrics
        for servername in excelDF['Server Name'].unique():
            metricsDF = excelDF[excelDF['Server Name'] == servername]
            metrics.saveMetrics(ServerName=servername, DirectEntitlements=len(metricsDF.query('(Type=="User")')))
            metrics.saveMetrics(ServerName=servername, GroupEntitlements=len(metricsDF.query('(Type!="User")')))

        # Group members sheet
        userDF = pd.read_sql_query('select u.UserID, u.DN, u.CostCentre, u.Name, u.EmployeeType, u.Company, u.Department, m.L0manager, m.L1manager, m.L2manager, m.L3manager, m.L4manager '
                                   'from users u left join userToManagers m on u.DN=m.DN ', conn1)

        groupMembersDF = pd.read_sql_query('select Name as SamAccountName, UserDN as DN from groupMembers', conn1)
        userGroupMergeDF = pd.merge(groupMembersDF, userDF, how='left', on='DN')
        excelDF = pd.merge(aclDF['SamAccountName'].drop_duplicates(), userGroupMergeDF, how='inner', on='SamAccountName')
        excelDF.drop(columns=['DN'], inplace=True)
        print('Length of Group members output sheet', len(excelDF))
        workbook.addSheet('Group members', excelDF)

        # Empty Groups
        emptyGroupsDF = pd.read_sql_query('select Name from groups where Name not in (select distinct Name from groupMembers)', conn1)
        emptyGroupsDF['SamAccountName'] = emptyGroupsDF['Name']
        excelDF = pd.merge(aclDF, emptyGroupsDF, how='inner', on='SamAccountName')
        print('Length of Empty groups output sheet', len(excelDF['Name'].drop_duplicates()))
        workbook.addSheet('Empty Groups', excelDF)

        workbook.save()



