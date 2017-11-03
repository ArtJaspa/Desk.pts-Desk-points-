#! /usr/bin/env python
scriptname = 'Desk.pts'
scriptversion = 'alpha'
scriptdesc = 'Project Desktop'
scriptauthor = 'Author: Art Jaspa (G46)'
scriptdate = '10/20/17'
scriptcopy = 'copyright: PTS-BDU, Innodata Knowledge Services, Inc.'

'''Revision History
10/18/17 alpha
'''

from tkinter import *
import pypyodbc, os, sqlite3 as lite, time, re, csv, codecs, subprocess, xml.etree.ElementTree as ET
from datetime import datetime, timedelta

global user
user = os.getenv('USERNAME')

class DeskGui():
    def __init__(self, master):
        self.master=master
        self.menubar=Menu(self.master)
        self.today=self.get_today()
        self.tomorrow=self.get_tomorrow()
        self.Gui()

    def TBC(self):
        filewin = Toplevel(self.master, width=50, height=50)
        button = Button(filewin, text="To be constructed")
        button.pack()


    def Gui(self):
        self.master.title('%s %s %s %s' % (project.upper(), scriptname, scriptversion, scriptdate))
        #self.master.maxsize(1420,400)
        self.master.maxsize(650,200)

        #maximize window
        m = self.master.maxsize()
        self.master.geometry('{}x{}+0+0'.format(*m))
        '''self.master.state = 'zoomed'#no effect
        self.master.attributes = ('-zooomed', True)#no effect'''
        
        self.logo = Canvas(self.master)
        self.logo.pack(fill='both')
        self.innologo = PhotoImage(file='innodata-isogen-logo.gif')
        self.logo.create_image(325, 100, anchor='center', image=self.innologo)
        
        self.TITOmenu = Menu(self.menubar, tearoff=0)
        self.TITOmenu.add_command(label="TITO", command=self.TITO)
        self.menubar.add_cascade(label="TITO", menu=self.TITOmenu)
        
        self.prodmenu = Menu(self.menubar, tearoff=0)
        self.prodmenu.add_command(label="User-Daily", command=self.userdailyprod)
        self.prodmenu.add_command(label="User-Weekly", command=self.userweeklyprod)
        self.prodmenu.add_command(label="Team-Daily", command=self.teamdailyprod)
        self.prodmenu.add_command(label="Team-Weekly", command=self.teamweeklyprod)
        self.prodmenu.add_command(label="Range", command=self.rangeprod)
        self.menubar.add_cascade(label="Productivity", menu=self.prodmenu)
        
        self.docmenu = Menu(self.menubar, tearoff=0)
        self.docmenu.add_command(label="Client Specs", command=self.specs)
        self.docmenu.add_command(label="Work Instruction", command=self.WI)
        self.docmenu.add_command(label="User Manual", command=self.usermanual)
        self.menubar.add_cascade(label="Documentation", menu=self.docmenu)

        self.querymenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Query", menu=self.querymenu)
        
        self.appmenu = Menu(self.menubar, tearoff=0)
        self.appmenu.add_command(label="Billing", command=self.billing)
        self.menubar.add_cascade(label="Application", menu=self.appmenu)
        
        self.exammenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Exam", menu=self.exammenu)
        
        self.sqlmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="SQL", menu=self.sqlmenu)
        
        self.bulletinmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Bulletin", menu=self.bulletinmenu)
        
        adminmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Admin", menu=adminmenu)
        
        self.optionsmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Options", menu=self.optionsmenu)
        
        self.helpmenu = Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="Help Index", command=self.TBC)
        self.helpmenu.add_command(label="About...", command=self.TBC)
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)

        self.master.config(menu=self.menubar)

    def TITO(self):
        self.TBC()
        return

    def prod(self):
        with codecs.open(self.tempout, 'w') as routput:
            outputwriter = csv.writer(routput,lineterminator='\n')
            outputwriter.writerow(['RefId', 'Content Name', 'BatchName', 'XML Filename', 'Task', 'Process Type', 'Productivity', 'uom', 'DateTime Start', 'DateTime End', 'Spent', 'Status', 'Remarks','UserId'])
            for a in self.TITOrecs:
                bs = self.batch_x[a[1]]
                b_split = bs.split('", "')
                b_split[0] = b_split[0][1:]
                b_split[-1] = b_split[-1][:-1]
                if b_split[0] == '' or b_split[0] == '-':
                    content_name = b_split[3]
                else:
                    content_name = b_split[0]
                xml_filename = b_split[1]
                process_type = b_split[2]
                outputwriter.writerow([a[0], content_name, a[1], xml_filename, initasklist_dict[str(a[7])], process_type, a[8], iniuomlist_dict[str(a[9])], a[3], a[4], a[5], inistatuslist_dict[str(a[6])], a[10], a[2]])
        return

    def time_to_timestamp(self,time2x):
        self.timestamp = datetime.fromtimestamp(time2x)
        a = self.timestamp.strftime('%m/%d/%y %H:%M:%S')
        return(a)

    def billing(self):
        print('%s - Started Billing' % self.time_to_timestamp(time.time()))
        command = 'Billing.pts.exe'
        subprocess.Popen(command, stdout=subprocess.PIPE, stdin=subprocess.PIPE, shell = True)
        print('%s - End' % self.time_to_timestamp(time.time()))
        return

    def userdailyprod(self):
        self.TBC()
        return

    def userweeklyprod(self):
        self.TBC()
        return

    def teamdailyprod(self):
        self.TBC()
        return
    
    def teamweeklyprod(self):
        self.TBC()
        return
    
    def teamprod(self):
        self.TBC()
        return

    def rangeprod(self):
        self.TBC()
        return

    def specs(self):
        self.TBC()
        return

    def WI(self):
        self.TBC()
        return

    def usermanual(self):
        self.TBC()
        return

    def reset(self):
        pass
        return

    def datetime_strftime_to_Guiformat(self, dts):
        return(dts.strftime('%m/%d/%y'))

    def get_today(self):
        self.today = time.strftime('%x',time.localtime())
        return(self.today)
    
    def get_tomorrow(self):
        self.tomorrow = self.datetime_strftime_to_Guiformat(datetime.fromtimestamp(time.time() + 24*3600))
        return(self.tomorrow)

    def get_week_start_end(self):
        day = self.get_today()
        dt = datetime.strptime(day, '%m/%d/%y')
        self.week_start = dt - timedelta(days=dt.weekday())
        self.week_end = self.week_start + timedelta(days=6)
        return
                          
    def connect_db(self):
        if dbsource == 'SQLserver':
            self.connection = pypyodbc.connect('Driver=SQL Server;' 'Server=10.160.0.21\VIRMDESQL21;' 'Database=OPUS_BLAW_V4;' 'uid=oos;pwd=it-oos')
        elif dbsource == 'SQLite':
            self.connection = lite.connect(sqlite_db)
        self.cursor = self.connection.cursor()
        return(self.connection)
       
    def get_from_db(self, curexecute):
        self.connect_db()
        if curexecute=='user_dailyprod':
            self.cursor.execute("SELECT RefId, BatchName, UserId, DateTimeStarted, DateTimeEnded, TimeSpent, StatusId, TaskId, Productivity, uom, Remarks from dbo.TITOpts_TaskInOut WHERE UserId='%s' AND DateTimeStarted>='%s' AND DateTimeEnded<'%s'" % (user, self.today, self.tomorrow))
        self.result=self.cursor.fetchall()
        self.connection.close()
        return(self.result)

    def update_db(self, curexecute):
        self.connect_db()
        if curexecute=='taskout':
            self.cursor.execute('UPDATE ' + self.TITOtable + ' SET DateTimeEnded=?, TimeSpent=?, Productivity=?, uom=?, StatusId=?, Remarks=? WHERE BatchName=? AND UserId=? AND TaskId=?', (self.te.get(), self.tsp.get(), self.prod.get(), self.uomvar.get().split(' ')[0], self.statusvar.get().split(' ')[0], self.remarkstext.get(), self.filename.get(), self.uservar.get(), self.taskvar.get().split(' ')[0]))
        if dbsource=='SQLServer':
            self.cursor.commit()
        else:
            self.connection.commit()
        self.connection.close()
        return

    def open_prodreport(self):
        if os.path.exists(self.tempout):
            filesize = os.stat(self.tempout).st_size
            if filesize == 134:
                print('No record match')
            elif filesize > 134:
                command = '"%s" "%s"' % (scalc, self.tempout)
                subprocess.Popen(command, stdout=subprocess.PIPE, stdin=subprocess.PIPE, shell = True)
        else:
            print('%s not found' % self.tempout)
        return
            
def process():
    root = Tk()
    DeskGui(root)
    root.mainloop()
    return()

def header():
    print('%s\n%s Python Script %s\n%s\n%s\n%s\n%s\n' % (project, scriptname, scriptversion, scriptdesc, scriptauthor, scriptdate, scriptcopy))
    return()

def main():
    cwd =(os.getcwd())
    inidb = '%s/%s.ini' % (cwd, scriptname)
    con = lite.connect(inidb)
    with con:
        cur = con.cursor()
        global project, scripttemp, sql_server, sql_db, sql_uid, sql_pwd, sqlite_db, scalc, dbsource, inistatuslist, inistatuslist_dict, initasklist, initasklist_dict, iniuomlist, iniuomlist_dict
        cur.execute('SELECT * from "configuration"')
        config = cur.fetchall()
        inivariables= {}
        for val, value, remarks in config:
            inivariables[val] = value
        dbsource = inivariables['dbsource']
        temp = inivariables['temp']
        usertemp = '%s/%s' % (temp, user)
        scripttemp = '%s/%s' % (usertemp,scriptname)
        tempfolders = (temp, usertemp, scripttemp)
        for a in tempfolders:
            if not(os.path.exists(a)):
                try:
                    os.mkdir(a)
                except:
                    print('Error in creating user temp folder "%s"' % a)
        tempout = '%s/%s/%s.csv' % (usertemp,scriptname,scriptname)
        if dbsource == 'SQLserver':
            sql_server = inivariables['sql_server']
            sql_db = inivariables['sql_db']
            sql_uid = inivariables['sql_uid']
            sql_pwd = inivariables['sql_pwd']
        project = inivariables['project']
        sqlite_db = inivariables['sqlite_db']
        scalc = inivariables['scalc']
        header()
    process()

if __name__ == '__main__':
    main()
