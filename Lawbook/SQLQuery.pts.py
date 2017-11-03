#! /usr/bin/env python
scriptname = 'SQLQuery.pts'
scriptversion = 'alpha'
scriptdesc = 'Execute Select command on MS SQL server'
scriptauthor = 'Author: Art Jaspa (G46)'
scriptdate = '10/12/17'
scriptcopy = 'copyright: PTS-BDU, Innodata Knowledge Services, Inc.'

from tkinter import *
import pypyodbc, os, sqlite3 as lite, time, re
from datetime import datetime, timedelta

global user
user = os.getenv('USERNAME')

class SQLGui():
    def __init__(self, master):
        self.master=master
        self.selentry=StringVar()
        self.today=self.get_today()
        self.tomorrow=self.get_tomorrow()
        self.selentry.set("SELECT RefId, BatchName, UserId, DateTimeStarted, DateTimeEnded, TimeSpent, StatusId, TaskId, Productivity, uom, Remarks from dbo.TITOpts_TaskInOut WHERE UserId='%s' AND DateTimeStarted>='%s' AND DateTimeEnded<'%s'" % (user, self.today, self.tomorrow))
        #self.selentry.set('')
        self.Gui()

    def Gui(self):
        self.master.title('%s %s %s' % (scriptname, scriptversion, scriptdate))
        self.master.maxsize(1400,50)

        self.selectentry=Entry(self.master, width=300, textvariable=self.selentry)
        self.selectentry.grid(row=0, sticky='nswe')
        self.selectentry.focus()

        self.OKbutton=Button(self.master, text='Execute', command=self.OK, width=10)
        self.OKbutton.grid(row=1, sticky='ns')

    def OK(self):
        curexecute = self.selentry.get()
        if curexecute.upper().startswith('SELECT '):
            self.connect_db()
            self.get_from_db(self.selentry.get())
            print(self.results)
        else:
            print("Error: Only 'SELECT command is allowed")
        return

    def datetime_strptime(self, dt):
        return(datetime.strptime(str(dt), "%Y-%m-%d %H:%M:%S.&f"))

    def datetime_strptime_to_Guiformat(self, dts):
        return(dts.strftime('%m/%d/%y'))

    def get_today(self):
        self.today = time.strftime('%x',time.localtime())
        return(self.today)
    
    def get_tomorrow(self):
        #self.tomorrow = self.datetime_strptime(datetime.fromtimestamp(time.time() + 24*3600))
        self.tomorrow = self.datetime_strptime_to_Guiformat(datetime.fromtimestamp(time.time() + 24*3600))
        return(self.tomorrow)
                          
    def connect_db(self):
        if dbsource == 'SQLserver':
            self.connection = pypyodbc.connect('Driver=SQL Server;' 'Server=10.160.0.21\VIRMDESQL21;' 'Database=OPUS_BLAW_V4;' 'uid=oos;pwd=it-oos')
        elif dbsource == 'SQLite':
            print(sqlite_db)
            self.connection = lite.connect(sqlite_db)
        self.cursor = self.connection.cursor()
        return(self.cursor)

    def get_from_db(self, curexecute):
        self.cursor.execute(curexecute)
        self.results = self.cursor.fetchone()
        self.connection.close()
        return(self.results)    

def process():
    root = Tk()
    frame = SQLGui(root)
    root.mainloop()
    return()

def header():
    print('%s Python Script %s\n%s\n%s\n%s\n%s\n' % (scriptname, scriptversion, scriptdesc, scriptauthor, scriptdate, scriptcopy))
    return()

def main():
    header()
    cwd =(os.getcwd())
    inidb = '%s/%s.ini' % (cwd, scriptname)
    con = lite.connect(inidb)
    with con:
        cur = con.cursor()
        global scalc, dbsource, sql_server, sql_db, sql_uid, sql_pwd, sqlite_db
        cur.execute('SELECT * from "configuration"')
        config = cur.fetchall()
        inivariables= {}
        for val, value, remarks in config:
            inivariables[val] = value
        dbsource = inivariables['dbsource']
        if dbsource == 'SQLserver':
            sql_server = inivariables['sql_server']
            sql_db = inivariables['sql_db']
            sql_uid = inivariables['sql_uid']
            sql_pwd = inivariables['sql_pwd']
        elif dbsource == 'SQLite':
            sqlite_db = inivariables['sqlite_db']
        scalc = inivariables['scalc']
    process()

if __name__ == '__main__':
    main()
