# Reads in scan files, sorts the entilements in to appropriate output files (no lines removed).
#
# Ouput files
# <servername>-local-groups.xlsx - entitlements associated with localgroups on the server defined in the localGroupsStartsWith list
# read-only-perms.xlsx - all entitlements that match those defined in the readOnlyPerms list
# everyone.xlsx - all entitlements that match those defined in the everyone list
# SID-perms.xlsx - perms with SID not user (these accounts are no longer active)
# group-user.csv - paths associated with groups or users - this will be used for active directory analysis in the WindowsAddUserDetails.py

import os
import csv
import pandas as pd
import Metrics
from datetime import datetime
from WriteExcel import Excel


def scan(dataPath, metrics: Metrics.TrackWindows):
    print('Processing', dataPath)
    dataOut = dataPath + '/results'  # directory for output file results

    # make output directory if it does not already exist
    if not os.path.isdir(dataOut):
        os.mkdir(dataOut)

    # patters used to direct rows to the correct output files
    localGroupsStartsWith = ('NT AUTHORITY\\', 'BUILTIN\\', 'CREATOR OWNER', 'CREATOR GROUP', 'NT SERVICE\\', 'APPLICATION PACKAGE AUTHORITY\\', 'IIS APPPOOL')
    everyone = ['US\\Domain Users', 'Everyone', 'NT AUTHORITY\\Authenticated Users']
    readOnlyPerms = ['ReadData', 'ReadPermissions', 'ReadAndExecute', 'Synchronize', 'Read', 'ExecuteFile', 'ReadExtendedAttributes', 'ReadAttributes']

    headers = ['Server Name', 'Folder Name', 'Group/User', 'Permissions', 'Applies To', 'Is Inherited', 'Top Level', 'Is top level', 'ACLs match top level']  # output file headers

    files = {}  # dictionary of output file details being produced
    fileRowCounts = {}  # file row counts for data processing checks
    fileServerRowCounts = {}  # file row counts for data processing checks
    perms = {}  # dictionary of root perms to evaluate when sub folders differ from the root

    # ############### Hepler functions #################

    def returnTopLevel(dirpath):
        # work out top level (root) folder from path. e.g. \\CRASDIABATTST03\e$\AutoRebJTJ\lib to \\CRASDIABATTST03\e$\AutoRebJTJ
        folders = dirpath.split('\\')
        if len(folders) >= 5:
            return '\\'.join(folders[0:5])
        return dirpath

    def buildTopLevelPerms(path, userGroup, perm):
        # If a a top level (root) folder is provided, add entitlements to perms dictionary
        topLevel = returnTopLevel(path)
        if path == topLevel:
            # top level so add to perms dictionary
            if topLevel in perms:
                if userGroup in perms[topLevel] and perm not in perms[topLevel][userGroup]:
                    perms[topLevel][userGroup].append(perm)
                else:
                    perms[topLevel][userGroup] = [perm]
            else:
                perms[topLevel] = {userGroup: [perm]}

    def compareTopLevelPerms(topLevel, userGroup, perm):
        # Returns Y if root level entitlements match those provided otherwise returns N
        if topLevel in perms and userGroup in perms[topLevel] and perm in perms[topLevel][userGroup]:
            return 'Y'
        else:
            return 'N'

    def allReadOnly(permissions):
        # Returns True if entitlements provide are all read only otherwise returns False
        for perm in permissions.split(', '):
            if perm not in readOnlyPerms:
                return False
        return True

    def saveMetrics(servername):
        # Saves metrics for the server provided
        metrics.saveMetrics(ServerName=servername,
                            LocalGroupsCount=fileServerRowCounts[(servername + '-local-groups.xlsx', servername)] if (servername + '-local-groups.xlsx', servername) in fileServerRowCounts else 0,
                            ReadOnlyCount=fileServerRowCounts[('read-only-perms.xlsx', servername)] if ('read-only-perms.xlsx', servername) in fileServerRowCounts else 0,
                            SIDcount=fileServerRowCounts[('SID-perms.xlsx', servername)] if ('SID-perms.xlsx', servername) in fileServerRowCounts else 0,
                            EveryoneCount=fileServerRowCounts[('everyone.xlsx', servername)] if ('everyone.xlsx', servername) in fileServerRowCounts else 0)

    def saveRow(filename, filepath, servername):
        # Saves data in filepath into appropriate structure with the files dictionary to be saved to file in saveFile()
        topLevel = returnTopLevel(filepath[1])
        match = compareTopLevelPerms(topLevel, filepath[2], filepath[3])
        isTopLevel = 'Y' if topLevel == filepath[1] else 'N'
        data = filepath + [topLevel, isTopLevel, match]
        if filename not in files:
            files[filename] = [data]
            fileRowCounts[filename] = 1
        else:
            files[filename].append(data)
            fileRowCounts[filename] += 1
        if (filename, servername) not in fileServerRowCounts:
            fileServerRowCounts[(filename, servername)] = 1
        else:
            fileServerRowCounts[(filename, servername)] += 1


    def saveFile(filename, header):
        # Save the provide file to disk. Split across multiple files/sheets to handle the 1 million row limit with excel
        print(filename, fileRowCounts[filename], datetime.now().strftime("%H:%M:%S"))
        increment = 1048570
        if filename.endswith('xlsx'):
            fromIndex = 0
            sheet = 1
            workbook = Excel(dataOut + '/' + filename)
            while fromIndex < len(files[filename]):
                toIndex = fromIndex + increment if (fromIndex + increment) < len(files[filename]) else len(files[filename])
                df = pd.DataFrame.from_records(files[filename][fromIndex:toIndex], columns=header)
                workbook.addSheet('Sheet ' + str(sheet), df)
                sheet += 1
                fromIndex += increment
            workbook.save()
            del workbook
        elif filename.endswith('csv'):
            fromIndex = 0
            filecount = 1
            while fromIndex < len(files[filename]):
                print('File', filecount)
                toIndex = fromIndex + increment if (fromIndex + increment) < len(files[filename]) else len(files[filename])
                outfile = open(dataOut + '/' + str(filecount) + '_' + filename, 'w', newline='', encoding='utf8')
                out = csv.writer(outfile)
                out.writerow(header)
                for data in files[filename][fromIndex:toIndex]:
                    out.writerow(data)
                filecount += 1
                fromIndex += increment
        else:
            print('Unknown file type', filename)
        files[filename] = []  # clear down the memory in files as this has now been saved to disk

    def appliesTo(str1):
        if str1 == 'None':
            return 'This folder only'
        elif str1 == 'ContainerInherit':
            return 'This folder and subfolders'
        elif str1 == 'ContainerInherit, ObjectInherit':
            return 'This folder, subfolders and files'
        elif str1 == 'ObjectInherit':
            return 'This folder and files'
        else:
            print('Unexpected applies to item', str1)
            exit(-1)

    # ################ Main #######################

    # Get list of files to process
    filelist = [f for f in os.listdir(dataPath) if os.path.isfile(os.path.join(dataPath, f))]

    totalRowCount = 0
    for f in filelist:  # Process every file in the source directory
        perms = {}  # clear perms list to reduce memory usage
        servername = f.upper().split('-ACL')[0]
        print('Processing', servername)
        # read through first time to populate top level perms
        fileRowCount = 0
        with open(dataPath + '/' + f, newline='', encoding='utf8') as csvfile:
            reader = csv.reader(csvfile)
            next(reader, None)  # skip header rows
            next(reader, None)  # skip header rows
            for row in reader:
                buildTopLevelPerms(row[0], row[1], row[2])
                fileRowCount += 1

        metrics.saveMetrics(ServerName=servername, ServerCsvRowCount=fileRowCount, ServerCsvDate=datetime.fromtimestamp(os.path.getmtime(dataPath + '/' + f)).strftime('%Y-%m-%d %H:%M:%S'))
        # read through populating output files
        with open(dataPath + '/' + f, newline='', encoding='utf8') as csvfile:
            reader = csv.reader(csvfile)
            next(reader, None)  # skip header rows
            next(reader, None)  # skip header rows
            for row in reader:
                # step through each line of the input file storing the data to the appropriate place in files dictionary
                totalRowCount += 1
                groupUser = row[1]
                permissions = row[2]
                rowDataComplete = True
                while len(row) < 5:  # Inheritance data is missing
                    row.append('N/A')
                    rowDataComplete = False
                if rowDataComplete: row[3] = appliesTo(row[3])
                if groupUser.startswith(localGroupsStartsWith) or groupUser.startswith(servername.upper()):
                    saveRow(servername + '-local-groups.xlsx', [servername] + row, servername)
                elif allReadOnly(permissions):
                    saveRow('read-only-perms.xlsx', [servername] + row, servername)
                elif groupUser in everyone:
                    saveRow('everyone.xlsx', [servername] + row, servername)
                elif groupUser.startswith('S-'):
                    saveRow('SID-perms.xlsx', [servername] + row, servername)
                else:
                    saveRow('group-user.csv', [servername] + row, servername)
        saveMetrics(servername)
        saveFile(servername + '-local-groups.xlsx', headers)  # save server specific file as no more data will be added to this

    # save files
    for filename in files:
        if not filename.endswith('-local-groups.xlsx'):
            saveFile(filename, headers)

    # do row count check to make sure we haven't dropped any between the input and output files
    sumRowCount = 0
    for f in fileRowCounts:
        sumRowCount += fileRowCounts[f]
    if totalRowCount != sumRowCount:
        print('Row count mismatch', totalRowCount, sumRowCount)
    else:
        print('Row count check good', totalRowCount, sumRowCount)
