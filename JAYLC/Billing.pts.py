#! /usr/bin/env python
scriptname = 'Billing.pts'
scriptversion = 'alpha rel. 2'
scriptdesc = 'Generate billing report from RC explaination html files'
scriptauthor = 'Author: Art Jaspa (G46)'
scriptdate = '10/23/17'
scriptcopy = 'copyright: PTS-BDU, Innodata Knowledge Services, Inc.'

'''Revision History
10/23/17 alpha rel. 2 1. Fix table not captured
2. change file format of rc from html to xml
3. have 

10/20/17 alpha
'''

from tkinter import *
import pypyodbc, os, sqlite3 as lite, time, re, codecs, xml.etree.ElementTree as ET, csv, subprocess
from datetime import datetime, timedelta
from os.path import join

try:
    import winreg
except ImportError:
    import _winreg as winreg
global user
user = os.getenv('USERNAME')

class BillingGui():
    def __init__(self, master):
        self.master = master
        if dbsource=='SQLite':
            self.shipmenttable='"%s"' % self.shipmenttable
        self.OKtitle='OK'
        self.pathvar = StringVar()
        self.pathvar.set('')
        self.pathstat=NORMAL
        self.getdefaultbg = 1
        self.errormsg=''
        self.Gui()

    def Gui(self):
        self.master.title('%s %s %s %s' % (project, scriptname, scriptversion, scriptdate))
        self.master.maxsize(800,80)

        m = self.master.maxsize()
        self.master.geometry('{}x{}+0+0'.format(*m))

        self.pathlabel = Label(self.master, text='Path: ')
        self.pathlabel.grid(row=0, column=0, sticky='e')
        self.path = Entry(self.master, text='%s' % self.pathvar, textvariable=self.pathvar, state=self.pathstat, width=125)
        self.path.grid(row=0, column=2, sticky='w')
        self.path.focus()
        
        self.defaultbg = self.path.cget('background')
        if self.getdefaultbg == 1:
            self.errorbg = self.errorfg = self.defaultbg
            self.getdefaultbg = 0
        
        commands = {'OK':self.OK}
        self.OK = Button(self.master, text=self.OKtitle, command=commands[self.OKtitle], width=10)
        self.OK.grid(row=2, column=2, sticky='e', padx=10, pady=10)
        
        self.pts = Label(self.master, text='PTS-BDU, Innodata', fg='blue')
        self.pts.grid(row=2, column=0, sticky='w', padx=10, pady=10, columnspan=3)

        self.error=Label(self.master, text='%s!' % self.errormsg, fg=self.errorfg, bg=self.errorbg)
        self.error.grid(row=3, column=0, columnspan=3, sticky='we', padx=10)

    def filedialog(self):#to be constructed
        self.pathvar.set(filedialog.askdirectory())
        self.Gui()
        return

    def filelisting(self):
        isempty = 1
        with open(listname, 'w') as filelist:
            for root, dirs, files in os.walk(self.pathvar.get()):
                for name in files:
                    filext = name.split('.')
                    filen = str(filext[0])
                    if str.upper(filext[-1]) == 'HTML' and str.upper(filext[-2]) == 'HTML-EXPLAINATION' and str.upper(filext[1]) == 'RC':
                        filelist.write('%s\n' % join(root, name))
        return

    def OK(self):
        self.filelisting()
        with open(listname, 'r') as filelist:
            self.i=0
            for self.file in filelist:
                self.file = re.sub(r'[\r\n\f]+','',self.file)
                self.parse_read()
        self.open_report()
        print('Done')
        return

    def get_from_db(self, curexecute):
        self.connect_db()
        if curexecute=='my_ongoing':
            self.cursor.execute('SELECT BatchName, DateTimeStarted, TaskId, Remarks, RefId FROM ' + self.TITOtable + ' WHERE UserId=? AND StatusId=3', (user,))
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
            
    def parse_read(self):
        filext = self.file.split('.')
        filen = str(filext[0])
        rcfile = '%s.rc.xml' % filen
        forparsing = {self.file, rcfile}
        self.iswellformed = 0
        for file in forparsing:
            with codecs.open(file, 'rb') as input:
                try:
                    self.tree = ET.parse(input)
                    self.iswellformed = 1
                except:
                    self.errormsg='Error: %s has parsing error!' % self.file
                    print(self.errormsg)
                    self.iswellformed = 0
                    break
        if self.iswellformed==1:
            changelist = {}
            with codecs.open(self.file, 'rb') as input2:
                self.tree = ET.parse(input2)
                for a in self.tree.iter():
                    if a.tag=='div' and 'class' in a.attrib and re.match(r'(modification|addition)', a.attrib['class']):
                        change = a.attrib['class']
                        path=a.findall('./*[@class="path"]')
                        for b in path:
                            changelist[b.text] = change
            with codecs.open(rcfile, 'rb') as input3:
                ET.register_namespace('','http://www.jpmorgan.com/olo')
                self.tree = ET.parse(input3)
                root = self.tree.getroot()
                if self.i == 0:
                    writemode = 'wb'
                else:
                    writemode = 'ab'
                with codecs.open(report, writemode, encoding='utf-8') as routput:
                    outputwriter = csv.writer(routput,lineterminator='\n')
                    if self.i == 0:
                        outputwriter.writerow(['SEQID', 'FILESIZE', 'TYPE', 'FILENAME', 'REMARKS', 'PATH'])
                    for a in changelist:
                        isfound=0
                        for b in self.tree.iter():
                            markup=self.rem_namespace(b.tag).upper()
                            if markup=='REFERENCEDCONTENT' and 'path' in b.attrib and b.attrib['path']==a:
                                self.i +=1
                                chunkfile='%s/%d.xml' % (scripttemp,self.i)
                                head = re.sub(r'^[\r\n\f]+','',b.tail)
                                btree = ET.ElementTree(b)
                                btreeroot = btree.getroot()
                                btree.write(tempout, encoding='utf-8')
                                with codecs.open(tempout, 'r', encoding='utf-8') as input4:
                                    with codecs.open(chunkfile, 'w', encoding='utf-8') as output:
                                        refcount = 0
                                        precount = 0
                                        tablecount = 0
                                        check_table = 0
                                        look_end_table = 0
                                        is_end_table = 0
                                        output.write(head)
                                        for c in input4:
                                            if re.match(r'^\s*<ReferencedContent',c):
                                                refcount = 1
                                            elif re.match(r'^\s*<Content',c):
                                                contentcount = 1
                                            elif re.match(r'^\s*<pre',c):
                                                precount = 1
                                            if refcount==1 and precount==1:
                                                if re.match(r'^.*</pre>\s*$',c):
                                                    check_table = 1
                                                    output.write(c)
                                                    continue
                                            output.write(c)
                                            if check_table==1 and look_end_table==0:
                                                if re.match(r'^\s*<table',c):
                                                    look_end_table = 1
                                                else:
                                                    break
                                            if look_end_table == 1:
                                                if re.match(r'^.*</table>\s*$',c):
                                                    is_end_table = 1
                                            if is_end_table == 1:
                                                break
                                isfound=1
                                outputwriter.writerow([self.i, os.stat(chunkfile).st_size, changelist[a], self.file, '', a])
                                break
                        if isfound==0:
                            outputwriter.writerow([self.i, '?', changelist[a], self.file, 'not found', a])
                            print('not found: %s' % (a))
        return
    
    def rem_namespace(self,s):
        s = s[s.find('}')+1:]
        return(s)

    def open_report(self):
        if os.path.exists(report):
            filesize = os.stat(report).st_size
            if filesize == 43:
                print('No record match')
            elif filesize > 43:
                command = '"%s" "%s"' % (scalc, report)
                subprocess.Popen(command, stdout=subprocess.PIPE, stdin=subprocess.PIPE, shell = True)
        else:
            print('%s not found' % report)
        return

def process():
    root = Tk()
    frame = BillingGui(root)
    root.mainloop()
    return()

def header():
    print('%s %s Python Script %s\n%s\n%s\n%s\n%s\n' % (project, scriptname, scriptversion, scriptdesc, scriptauthor, scriptdate, scriptcopy))
    return()

def main():
    cwd =(os.getcwd())
    inidb = '%s/%s.ini' % (cwd, scriptname)
    con = lite.connect(inidb)
    with con:
        cur = con.cursor()
        global scalc, dbsource, sql_server, sql_db, sql_uid, sql_pwd, sqlite_db, project, tempout, scripttemp, listname, report
        cur.execute('SELECT * from "configuration"')
        config = cur.fetchall()
        inivariables= {}
        for val, value, remarks in config:
            inivariables[val] = value
        dbsource = inivariables['dbsource']
        project = inivariables['project']
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
        listname = '%s/%s/%s.lst' % (usertemp,scriptname,scriptname)
        tempout = '%s/%s/%s.xml' % (usertemp,scriptname,scriptname)
        report = '%s/%s/%s.csv' % (usertemp,scriptname,scriptname)
        if os.path.exists(report):
            os.remove(report)
        if dbsource == 'SQLserver':
            sql_server = inivariables['sql_server']
            sql_db = inivariables['sql_db']
            sql_uid = inivariables['sql_uid']
            sql_pwd = inivariables['sql_pwd']
        project = inivariables['project']
        sqlite_db = inivariables['sqlite_db']
        scalc = inivariables['scalc']
        scalcsplit = scalc.split('; ')
        for a in scalcsplit:
            if os.path.exists(a):
                scalc=a
                break
        program = {'excel.exe', 'scalc.exe'}
        regpath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
        for a in program:
            try:
                handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, '%s\%s' % (regpath, a))
                num_values = winreg.QueryInfoKey(handle)[1]
                for i in range(num_values):
                    j = winreg.EnumValue(handle, i)
                    if j[0]=='':
                        scalc=j[1]
                break
            except:
                pass
    header()
    process()

if __name__ == '__main__':
    main()
