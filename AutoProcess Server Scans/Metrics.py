# Class to handle metrics
import os
import sqlite3
import pandas as pd
from datetime import datetime
from WriteExcel import Excel

MetricsDB = 'Metrics.db'


class TrackWindows:
    def __init__(self, path, env):
        # Work out scan date from input files
        self.Env = env
        filelist = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        sourceFileCount = 0
        dateSum = 0
        for f in filelist:
            dateSum += os.path.getmtime(path + '/' + f)
            sourceFileCount += 1
        self.ScanDate = datetime.fromtimestamp(int(dateSum/sourceFileCount)).strftime('%Y-%m-%d')
        print('Scan Date:', self.ScanDate)
        conn = sqlite3.connect(MetricsDB)
        conn.cursor().execute('update windowsACL set State="Old" where ScanDate=? and Environment=?', (self.ScanDate, env))
        conn.commit()
        conn.close()

    def saveMetrics(self, ServerName, ServerCsvDate=0, ServerCsvRowCount=0, LocalGroupsCount=0, ReadOnlyCount=0, SIDcount=0, EveryoneCount=0, DirectEntitlements=0, GroupEntitlements=0):
        conn = sqlite3.connect(MetricsDB)
        c = conn.cursor()
        for (CsvDate, CsvRowCount, LocalGroups, ReadOnly, SID, Everyone, DirectPerms, GroupPerms) in c.execute('select ServerCsvDate, ServerCsvRowCount, LocalGroupsCount, ReadOnlyCount, SIDcount, EveryoneCount, DirectEntitlements, '
                                                                                                               'GroupEntitlements from windowsACL where ScanDate=? and Environment=? and ServerName=? and State="Current"',
                                                                                                               (self.ScanDate, self.Env, ServerName)):
            ServerCsvDate = CsvDate if ServerCsvDate == 0 else ServerCsvDate
            ServerCsvRowCount += CsvRowCount
            LocalGroupsCount += LocalGroups
            ReadOnlyCount += ReadOnly
            SIDcount += SID
            EveryoneCount += Everyone
            DirectEntitlements += DirectPerms
            GroupEntitlements += GroupPerms

        c.execute('delete from windowsACL where ScanDate=? and Environment=? and ServerName=?', (self.ScanDate, self.Env, ServerName))
        Total = LocalGroupsCount + ReadOnlyCount + SIDcount + EveryoneCount + DirectEntitlements + GroupEntitlements
        Mismatch = ServerCsvRowCount - Total
        c.execute('insert into windowsACL values (?,?,?,?,?,?,?,?,?,?,?,?,?,"Current")', (self.ScanDate, self.Env, ServerName, ServerCsvDate, ServerCsvRowCount, LocalGroupsCount, ReadOnlyCount, SIDcount, EveryoneCount,
                                                                                          DirectEntitlements, GroupEntitlements, Total, Mismatch))
        conn.commit()
        conn.close()

    def writeExcel(self, path):
        workbook = Excel(path + '/results/Metrics.xlsx')
        conn = sqlite3.connect(MetricsDB)
        excelDF = pd.read_sql_query('select ScanDate as "Scan Date", Environment, ServerName as "Server Name", ServerCsvDate as "Server CSV Timestamp", ServerCsvRowCount as "Source Scan file (<servername>-ACL-FoldersOnly.csv)", '
                                    'LocalGroupsCount as "<servername>-local-groups.xlsx", ReadOnlyCount as "read-only-perms.xlsx", SIDcount as "SID-perms.xlsx ", EveryoneCount as "everyone.xlsx", '
                                    'DirectEntitlements as "group-user with AD info.xlsx:Users", GroupEntitlements as "group-user with AD info.xlsx:Groups", Total, Mismatch as "Source vs Results Mismatch count" '
                                    'from windowsACL where ScanDate=? and Environment=?', conn, params=(self.ScanDate, self.Env))
        conn.close()
        workbook.addSheet(self.Env, excelDF)
        workbook.save()


class TrackUnix:
    def __init__(self, path, env):
        # Work out scan date from input files
        self.Env = env
        filelist = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        sourceFileCount = 0
        dateSum = 0
        for f in filelist:
            dateSum += os.path.getmtime(path + '/' + f)
            sourceFileCount += 1
        self.ScanDate = datetime.fromtimestamp(int(dateSum/sourceFileCount)).strftime('%Y-%m-%d')
        print('Scan Date:', self.ScanDate)
        conn = sqlite3.connect(MetricsDB)
        conn.cursor().execute('update unixACL set State="Old" where ScanDate=? and Environment=?', (self.ScanDate, env))
        conn.commit()
        conn.close()

    def saveMetrics(self, ServerName, ServerTxtDate, EntitlementCount, ServerAccess):
        conn = sqlite3.connect(MetricsDB)
        c = conn.cursor()
        c.execute('delete from unixACL where ScanDate=? and Environment=? and ServerName=?', (self.ScanDate, self.Env, ServerName))
        c.execute('insert into unixACL values (?,?,?,?,?,?,"Current")', (self.ScanDate, self.Env, ServerName, ServerTxtDate, EntitlementCount, ServerAccess))
        conn.commit()
        conn.close()

    def writeExcel(self, path):
        workbook = Excel(path + '\\Metrics.xlsx')
        conn = sqlite3.connect(MetricsDB)
        excelDF = pd.read_sql_query('select ScanDate as "Scan Date", Environment, ServerName as "Server Name", ServerTxtDate as "Server txt file Timestamp", EntitlementCount as "Server Entitlement Count", '
                                    'ServerAccess as "Server Logon Count" from unixACL where ScanDate=? and Environment=?', conn, params=(self.ScanDate, self.Env))
        conn.close()
        workbook.addSheet(self.Env, excelDF)
        workbook.save()
