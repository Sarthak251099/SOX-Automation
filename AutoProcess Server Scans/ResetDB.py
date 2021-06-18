import sqlite3
import os

def resetdb(env):
    dbFile = env + '_LDAPcache.db'
    drop = False
    if os.path.isfile(dbFile):
        print('True')
        drop = True

    conn = sqlite3.connect(dbFile)
    c = conn.cursor()

    if drop:
        c.execute('drop table users')
        c.execute('drop table userToManagers')
        c.execute('drop table groups')
        c.execute('drop table groupMembers')
        conn.commit()

    c.execute(
        "CREATE TABLE users (DN varchar(100) primary key, UserID varchar(100), ManagerDN varchar(100), CostCentre varchar(100), Name varchar(100), EmployeeType varchar(100), Company varchar(100), Department varchar(100), insertDate varchar(30))")
    c.execute("CREATE TABLE userToManagers (DN varchar(100) primary key, L0manager varchar(100), L1manager varchar(100), L2manager varchar(100), L3manager varchar(100), L4manager varchar(100))")
    c.execute("CREATE TABLE groups (Name varchar(100), DN varchar(100), Description text, Notes text, insertDate varchar(30), primary key(Name, DN))")
    c.execute("CREATE TABLE groupMembers (Name varchar(100), UserDN varchar(100), primary key(Name, UserDN))")

    conn.commit()

    c.execute("create index usersUserIDindex on users(UserID)")

    conn.commit()

    c.execute("vacuum")

    conn.commit()


