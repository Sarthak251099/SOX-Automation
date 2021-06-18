# This provides all the functions to look up groups and users in active directory with the results saved into LDAPcache.db for performance
# to use for producing reports.
# Active directory is comparatively slow for the amount of look-ups required so code is structured to minimise this in all cases
# Usage:
#   userInDB(userID) - returns True if userID in cache, false otherwise
#   groupInDB(groupName) - returns True if group in cache, false otherwise
#   populateUserFromAD(userID) - populates <env>_LDAPcache.dp from active directory and returns true is UserID found otherwise returns false
#   populateGroupFromAD(groupName) - populates <env>_LDAPcache.dp from active directory and returns true is group found otherwise returns false

import pyad.aduser
import sqlite3

topAccounts = ['mmullin', 'mbloom', 'hwaters1', 'wfuller'] # stop looking up the tree when get to these people. 'hwaters1', 'wfuller' added as they are listed as reporting to each other??? so can never get any further!
forceManagerUpdate = False  # change to True to for manager updates if have changed topAccounts.


class Populate:
    def __init__(self, env):
        # connect to db
        self.conn = sqlite3.connect(env + '_LDAPcache.db')
        self.cache = self.conn.cursor()

        if forceManagerUpdate:
            self.cache.execute('delete from userToManagers')
            self.conn.commit()

        # connect to active directory (ldap)
        self.ldapUser = pyad.adquery.ADQuery()
        self.ldapGroup = pyad.adquery.ADQuery()

        # In memory dictionaries so only go back to active directory once per userid/group
        self.users = {}
        self.userDNtoID = {}
        self.userToManagers = {}
        self.groups = {}

        self.groupsNotFound = []  # build list of groups not found so only query ldap once
        self.usersNotFound = []  # build list of users not found so only query ldap once
        self.usersDnNotFound = []  # build list of users not found so only query ldap once

        # Load data from DB and populates users, userDNtoID, usersToManagers, groups, groupMembers
        for (DN, UserID, ManagerDN, CostCentre, Name, EmployeeType, Company, Department) in self.cache.execute('select DN, UserID, ManagerDN, CostCentre, Name, EmployeeType, Company, Department from users'):
            self.users[UserID.lower()] = dict(DN=DN, UserID=UserID, ManagerDN=ManagerDN, CostCentre=CostCentre, Name=Name, EmployeeType=EmployeeType, Company=Company, Department=Department)
            self.userDNtoID[DN] = UserID.lower()

        for (UserID, DN, L0manager, L1manager, L2manager, L3manager, L4manager) in self.cache.execute('select u.UserID, m.DN, m.L0manager, m.L1manager, m.L2manager, m.L3manager, m.L4manager from users u, userToManagers m where u.DN=m.DN'):
            self.userToManagers[UserID.lower()] = dict(DN=DN, L0manager=L0manager, L1manager=L1manager, L2manager=L2manager, L3manager=L3manager, L4manager=L4manager)

        for (Name, DN, Description, Notes) in self.cache.execute('select Name, DN, Description, Notes from groups'):
            self.groups[Name] = dict(DN=DN, Description=Description, Notes=Notes)

    def userInDB(self, UserID):
        return True if UserID.lower() in self.users else False

    def groupInDB(self, Name):
        return True if Name in self.groups else False

    # ##### User  lookup functions

    def __insertIntoUserCache(self, lr):
        # Inserts record provided to LDAPcache.db and to in memory dictionaries
        print('Cache miss (user)', lr)
        self.cache.execute('insert into users values (?,?,?,?,?,?,?,?,datetime("now"))', (lr['distinguishedName'], lr['SamAccountName'], lr['Manager'], lr['extensionAttribute6'], lr['DisplayName'], lr['EmployeeType'], lr['Company'], lr['Department']))
        self.conn.commit()
        self.users[lr['SamAccountName'].lower()] = dict(DN=lr['distinguishedName'], UserID=lr['SamAccountName'], ManagerDN=lr['Manager'], CostCentre=lr['extensionAttribute6'], Name=lr['DisplayName'], EmployeeType=lr['EmployeeType'],
                                                   Company=lr['Company'], Department=lr['Department'])
        self.userDNtoID[lr['distinguishedName']] = lr['SamAccountName'].lower()
        return True

    def __insertManagersIntoDB(self, lr):
        # Inserts record provided to LDAPcache.db and to in memory dictionaries
        #print('Inserting Manager', lr)
        self.cache.execute('insert into userToManagers values (?,?,?,?,?,?)', (lr['DN'], lr['L0manager'], lr['L1manager'], lr['L2manager'], lr['L3manager'], lr['L4manager']))
        self.conn.commit()
        self.userToManagers[lr['UserID'].lower()] = dict(DN=lr['DN'], L0manager=lr['L0manager'], L1manager=lr['L1manager'], L2manager=lr['L2manager'], L3manager=lr['L3manager'], L4manager=lr['L4manager'])
        return True


    def __ADqueryUserByIDinsertDB(self, userID):
        # Queries AD by userID. If found insert into DB and users/usersTNtoID dict, return True. If not found return False.
        if userID.lower() in self.users:
            return True
        elif userID in self.usersNotFound:
            return False
        self.ldapUser.execute_query(attributes=['SamAccountName','distinguishedName','Manager','extensionAttribute6','DisplayName','EmployeeType', 'Company', 'Department'],
                           where_clause="objectCategory='user' and SamAccountName='" + userID + "'",
                           base_dn="DC=us,DC=aegon,DC=com")
        for row in self.ldapUser.get_results():
            return self.__insertIntoUserCache(row)
        self.usersNotFound.append(userID)
        return False


    def __ADqueryUserByDNinsertDB(self, userDN):
        # Queries AD by userDN. If found insert into DB and users/usersTNtoID dict, return True. If not found return False.
        if userDN in self.userDNtoID:
            return True
        elif "'" in userDN:  # Can't handle O'Brien type names currently
            return False
        elif userDN in self.usersDnNotFound:
            return False
        self.ldapUser.execute_query(attributes=['SamAccountName','distinguishedName','Manager','extensionAttribute6','DisplayName','EmployeeType', 'Company', 'Department'],
                        where_clause="distinguishedName='" + userDN + "'",
                        base_dn="DC=us,DC=aegon,DC=com")
        for row in self.ldapUser.get_results():
            return self.__insertIntoUserCache(row)
        self.usersDnNotFound.append(userDN)
        return False


    def populateUserFromAD(self, UserID):
        # populates DB and users/usersTNtoID dict
        # returns true if found in LDAP, false if not
        if UserID.lower() in self.users and UserID.lower() in self.userToManagers:
            # UserID already in DB so nothing to do
            return True

        if self.__ADqueryUserByIDinsertDB(UserID):
            # Populate Manager details into DB
            userDetails = self.users[UserID.lower()].copy()

            # for ADM and DAD admin accounts use the person not admin account for manager lookups as considered more accurate
            if userDetails['EmployeeType'] == 'ADM' or userDetails['EmployeeType'] == 'DAD':
                admID = UserID[3:].lower() if userDetails['EmployeeType'] == 'ADM' else UserID[4:].lower()
                if self.__ADqueryUserByIDinsertDB(admID):
                    resultsADM = self.users[admID]
                    if resultsADM and 'ManagerDN' in resultsADM and resultsADM['ManagerDN'] and 'Company' in resultsADM:
                        print('ADM - Using normal account')
                        userDetails['ManagerDN'] = resultsADM['ManagerDN']
                        userDetails['Company'] = resultsADM['Company']

            results = {'DN': userDetails['DN'],
                       'UserID': UserID}

            # iterate up through manager tree
            managers = []
            while 'ManagerDN' in userDetails and userDetails['ManagerDN'] and userDetails['UserID'] not in topAccounts and userDetails['ManagerDN'] != userDetails['DN'] and self.__ADqueryUserByDNinsertDB(userDetails['ManagerDN']) and len(managers) < 10:
                managerDetails = self.users[self.userDNtoID[userDetails['ManagerDN']]]
                if managerDetails and 'Name' in managerDetails:
                    #print('Manager details:', managerDetails, len(managers))
                    managers.append(managerDetails['Name'])
                userDetails = managerDetails

            results['L0manager'] = ''
            results['L1manager'] = ''
            results['L2manager'] = ''
            results['L3manager'] = ''
            results['L4manager'] = ''

            # populate manger tree depending on how many manager levels were found
            if len(managers) == 0:
                pass
            elif len(managers) == 1:
                results['L0manager'] = managers[0]
            elif len(managers) == 2:
                results['L0manager'] = managers[0]
                results['L3manager'] = managers[-2]
                results['L4manager'] = managers[-1]
            elif len(managers) == 3:
                results['L0manager'] = managers[0]
                results['L2manager'] = managers[-3]
                results['L3manager'] = managers[-2]
                results['L4manager'] = managers[-1]
            else:
                results['L0manager'] = managers[0]
                results['L1manager'] = managers[-4]
                results['L2manager'] = managers[-3]
                results['L3manager'] = managers[-2]
                results['L4manager'] = managers[-1]
            return self.__insertManagersIntoDB(results)

        else:  # User ID not found in AD
            return False


    # ##### Group functions

    def __insertIntoGroupCache(self, name, lr):
        # Inserts record provided to LDAPcache.db and to in memory dictionaries
        print('Cache miss (group DN)', name, lr)
        description = ';'.join(lr['Description']) if lr['Description'] else None
        self.cache.execute('insert into groups values (?,?,?,?,datetime("now"))', (name, lr['distinguishedName'], description, lr['Info']))
        self.conn.commit()
        self.groups[name] = dict(DN=lr['distinguishedName'], Description=description, Notes=lr['Info'])
        return True

    def __ADqueryGroupInsertDB(self, Name):
        # Queries AD by Name. If found insert into DB and groups dict, return True. If not found return False.
        if Name in self.groupsNotFound:
            return False
        self.ldapGroup.execute_query(attributes=['distinguishedName', 'Description', 'Info'],
                           where_clause="objectCategory='CN=Group,CN=Schema,CN=Configuration,DC=us,DC=aegon,DC=com' and SamAccountName='" + Name + "'",
                           base_dn="DC=us,DC=aegon,DC=com")
        for row in self.ldapGroup.get_results():
            return self.__insertIntoGroupCache(Name, row)
        self.groupsNotFound.append(Name)
        return False

    def __populateGroupMembersFromLdap(self, Name):
        # Inserts record provided to LDAPcache.db and to in memory dictionaries
        groupDN = self.groups[Name]['DN']
        self.ldapGroup.execute_query(attributes=['distinguishedName', 'SamAccountName'],
                        where_clause="objectCategory='user' and memberOf='" + groupDN + "'",
                        base_dn="DC=us,DC=aegon,DC=com")
        for row in self.ldapGroup.get_results():
            self.cache.execute('insert into groupMembers values (?,?)', (Name, row['distinguishedName']))
            self.conn.commit()
            self.populateUserFromAD(row['SamAccountName'])
        return True

    def populateGroupFromAD(self, Name):
        # populates DB and users/usersTNtoID dict. returns true if found in LDAP, false if not
        if Name in self.groups:
            # Group is already in DB so nothing to do
            return True
        if self.__ADqueryGroupInsertDB(Name):
            # Get groups members
            return self.__populateGroupMembersFromLdap(Name)
        else:
            return False

