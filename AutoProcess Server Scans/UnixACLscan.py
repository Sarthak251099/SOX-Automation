# Reads in scan files, sorts the entilements in to appropriate sheets (no lines removed).
# User IDs are resolved with manger details using active directory via GetUserGroup module

import os
import pandas as pd
import sqlite3
from datetime import datetime
from WriteExcel import Excel
import GetUserGroup
import ResetDB
import Metrics

def serverScanLinux(path):
    env = 'LinuxProd'
    OutputFile = path + '\\ScanResults.xlsx'  # results file

    # excel sheet headers
    ownerPermHeaders = ['Server Name', 'Path', 'UserID', 'Entitlement']
    permHeaders = ['Server Name', 'Path', 'Group/User', 'Entitlement']
    logonHeaders = ['Server Name', 'Account Name', 'UserID']

    # setup GetUserGroup / ldapCache.db
    ResetDB.resetdb(env)
    ad = GetUserGroup.Populate(env)

    # Initialise metrics
    metrics = Metrics.TrackUnix(path=path, env=env)

    # connect to ldap cache db
    conn1 = sqlite3.connect(env + '_LDAPcache.db')
    cache1 = conn1.cursor()


    # ############## Helper functions #####################

    def permTranslate(perm):
        # translates unix style rwx perms into read, write execute.
        results = []
        if perm[0] == 'r':
            results.append('Read')
        if perm[1] == 'w':
            results.append('Write')
        if perm[2] != '-':
            results.append('Execute')
        return ', '.join(results)


    def removeDomain(str1):
        # Strips off US or DS domain name from user/group
        if str1.startswith('US\\') or str1.startswith('DS\\'):
            return str1.split('\\')[1]
        else:
            print('Unrecognised Domain', str1)
            exit(-1)


    # ################# Main ########################

    # Generate the list of input files
    fileList = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and (f.endswith('.out') or f.endswith('.txt'))]

    ownerPerms = []
    groupPerms = []
    worldPerms = []
    logon = []
    accounts = {}

    print(','.join(['Server Name', 'Server Entitlement Count', 'Server Logon Count']))

    totalRowCount = 0
    for f in fileList:  # Process every file in the source directory
        #print(f)
        servername = f.lower().split('.')[0]
        serverPermCount = 0
        serverLogonCount = 0
        with open(path + '/' + f, newline='') as inFile:
            foundPermStart = False
            foundLogonStart = False
            Done = False
            rowcount = 0
            for line in inFile:
                rowcount += 1
                line = line.rstrip()  # remove trailing spaces and newline characters
                if foundLogonStart and len(line) == 0:
                    # reached end of login section
                    Done = True
                elif Done or len(line) == 0:
                    pass  # empty line, ignore
                elif line.startswith('Permissions'):  # detect start of Permissions section of file
                    foundPermStart = True
                elif not foundLogonStart and line.startswith('Users allowed to login'):  # detect start of users with logon access
                    foundPermStart = False
                    foundLogonStart = True
                elif line.startswith('Ending scan') or line.startswith('Executing on') or line.startswith('Starting scan') or line.startswith('List of directories') or line.startswith('The operating system is') \
                        or line.startswith('Application folders to search are') or line.startswith('WARNING: This list contains only locally-cached users') or line.startswith('Determine Operating System'):
                    pass  # lines with text but do not contain information we need
                elif line.startswith('/') and not foundPermStart and not foundLogonStart :
                    pass  # filter out folder list under 'Application folders to search are:' at start of file as this is not needed
                elif foundPermStart and line.startswith('d'):  # Line has permission information
                    # Get owner level entitlements
                    data = line.split(' ', 3)
                    perms = permTranslate(data[0][1:4])
                    ownerPerms.append([servername, data[3], data[1], perms])
                    # Get group level entitlements
                    accounts[data[1]] = 1
                    perms = permTranslate(data[0][4:7])
                    groupPerms.append([servername, data[3], data[2], perms])
                    # Get world level entitlements
                    perms = permTranslate(data[0][7:10])
                    worldPerms.append([servername, data[3], '-', perms])
                    totalRowCount += 1
                    serverPermCount += 1
                elif foundLogonStart and (line.startswith('US') or line.startswith('DS')):  # Line with logon access data
                    (account, VASdetail) = line.split(':', 1)
                    SAMaccount = removeDomain(account)
                    accounts[SAMaccount] = 1  # add to accounts so can check it is in the ldap cache db later
                    logon.append([servername, account, SAMaccount])
                    serverLogonCount += 1
                elif foundLogonStart:
                    print('Warning, non-domain login:', line)
                else:
                    print('Unhandled row:', line, ';')
                    print(rowcount, foundLogonStart, foundPermStart, Done)
                    exit(-1)
        print(','.join([servername, str(serverPermCount), str(serverLogonCount)]))
        metrics.saveMetrics(ServerName=servername, ServerTxtDate=datetime.fromtimestamp(os.path.getmtime(path + '/' + f)).strftime('%Y-%m-%d %H:%M:%S'), EntitlementCount=serverPermCount, ServerAccess=serverLogonCount)

    # Ensure ldap cache has a record for every account by iterating through each item in accounts
    print()
    print('Starting AD lookups')
    rowcount = 0
    notFound = []
    for SAMaccount in accounts.keys():
        rowcount += 1
        if not ad.populateUserFromAD(SAMaccount):
            if SAMaccount not in notFound:
                notFound.append(SAMaccount)
        if rowcount % 100 == 0:
            print(rowcount, datetime.now().strftime("%H:%M:%S"))
    print('Account records not found in Active Directory :', len(notFound), ':', notFound)

    print()
    # Do row count check to make sure input row count matches output row count to catch any unexpected issues
    print('Expecting', totalRowCount, 'for each entitlement sheet')

    # Produce Excel file from combination of the parsed input files and contents of the ldap cache db (active directory)
    workbook = Excel(OutputFile)
    userDF = pd.read_sql_query('select u.UserID, u.CostCentre, u.Name, u.EmployeeType, u.Company, m.L0manager, m.L1manager, m.L2manager, m.L3manager, m.L4manager '
                            'from users u left join userToManagers m on u.DN=m.DN ', conn1)

    # Owner entitlements
    ownerDF = pd.DataFrame.from_records(ownerPerms, columns=ownerPermHeaders)
    ownerExcelDF = pd.merge(ownerDF, userDF, how='left', on='UserID')
    workbook.addSheet('Owner entitlements', ownerExcelDF)
    print('Length of Owner output sheet', len(ownerExcelDF))

    # Group entitlements
    groupDF = pd.DataFrame.from_records(groupPerms, columns=permHeaders)
    workbook.addSheet('Group entitlements', groupDF)
    print('Length of Group output sheet', len(groupDF))

    # World entitlements
    worldDF = pd.DataFrame.from_records(worldPerms, columns=permHeaders)
    workbook.addSheet('World entitlements', worldDF)
    print('Length of World output sheet', len(worldDF))

    # Server logon access
    logonDF = pd.DataFrame.from_records(logon, columns=logonHeaders)
    logonExcelDF = pd.merge(logonDF, userDF, how='left', on='UserID')
    workbook.addSheet('Server logon access', logonExcelDF)
    print('Length of Logon output sheet', len(logonExcelDF))

    workbook.save()

    metrics.writeExcel(path=path)