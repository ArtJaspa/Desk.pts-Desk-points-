#! /usr/bin/env python
scriptname = 'TITO.pts'
scriptversion = 'alpha'
scriptdesc = 'TaskIn/TaskOut'
scriptauthor = 'Author: Art Jaspa (G46)'
scriptdate = '11/2/17'
scriptcopy = 'copyright: PTS-BDU, Innodata Knowledge Services, Inc.'

'''Revision History
11/2/17 alpha
'''

from tkinter import *
import pypyodbc, os, sqlite3 as lite, time, re
from tkinter import ttk
from datetime import datetime, timedelta
from tkinter import filedialog

global user
user = os.getenv('USERNAME')

class TITOGui():
    def __init__(self, master):
        self.TITOtable = 'TITOpts_TaskInOut'
        self.Usertable = 'dbo.TITOpts_User'
        if dbsource=='SQLite':
            self.shipmenttable='"%s"' % self.shipmenttable
            self.TITOtable='"%s"' % self.TITOtable
            self.Usertable='"%s"' % self.Usertable
        self.OKtitle='Taskin'
        self.timestart = self.timeend = self.timespent = self.filename = ''
        self.batchnamestat=NORMAL
        self.master = master
        self.filename = StringVar()
        self.filename.set('')
        self.RefIdvar = StringVar()
        self.RefIdvar.set('')
        self.RefIdstat=DISABLED
        self.ClusterIdvar = StringVar()
        self.ClusterIdvar.set('')
        self.ClusterIdstat=DISABLED
        self.ts = StringVar()
        self.ts.set('')
        self.tsstat=DISABLED
        self.te = StringVar()
        self.te.set('')
        self.testat=DISABLED
        self.tsp = StringVar()
        self.tsp.set('')
        self.tspstat=DISABLED
        self.editvar = IntVar()
        self.editvar.set(0)
        self.prod = StringVar()
        self.prod.set('')
        self.prodstat=DISABLED
        self.uomlist = uomlist
        self.uomvar = StringVar()
        self.uomvar.set('NULL')
        self.tasklist = initasklist
        self.taskvar = StringVar()
        self.taskvar.set('NULL')
        self.statuslist = inistatuslist
        self.statusvar = StringVar()
        self.statusvar.set('NULL')
        self.uomstat=DISABLED
        self.statusstat=DISABLED
        self.remarksstat=NORMAL
        self.taskstat=NORMAL
        self.remarkstext = StringVar()
        self.remarkstext.set('')
        self.getdefaultbg = 1
        self.checkongoing = 1
        self.userstat=DISABLED
        self.uservar = StringVar()
        self.uservar.set(user)
        self.editstat=DISABLED
        self.errormsg=''
        
        self.master.title('%s %s %s' % (scriptname, scriptversion, scriptdate))
        self.master.maxsize(340,500)

        #m = self.master.maxsize()
        #self.master.geometry('{}x{}+0+0'.format(*m))

        self.get_from_db('is_edit_user')
        '''#deactivated temporarily
        if len(self.result)>0:
            self.editstat=NORMAL'''
        self.editlabel = Label(self.master, text='Edit?')
        self.editlabel.grid(row=0, column=0, sticky='e')
        self.edit = Checkbutton(self.master, variable=self.editvar, state=self.editstat, command=self.edit_init)
        self.edit.grid(row=0, column=2, sticky='w')
        #self.edit.bind('<Button-1>', self.edit_init)

        self.projectlabel = Label(self.master, text='Project: ')
        self.projectlabel.grid(row=1, column=0, sticky='e')
        self.project = Label(self.master, text=project)
        self.project.grid(row=1, column=2, sticky='w')

        self.userlabel = Label(self.master, text='UserID: ')
        self.userlabel.grid(row=2, column=0, sticky='e')
        self.user = Entry(self.master, text='%s' % self.uservar, textvariable=self.uservar, state=self.userstat, width=5)
        self.user.grid(row=2, column=2, sticky='w')
        self.defaultbg = self.userlabel.cget('background')
        if self.getdefaultbg == 1:
            self.errorbg = self.errorfg = self.defaultbg
            self.getdefaultbg = 0

        self.tasklabel = Label(self.master, text='Task: ')
        self.tasklabel.grid(row=3,sticky='e')
        self.task = OptionMenu(self.master,self.taskvar,*self.tasklist)
        self.task.config(state=self.taskstat)
        self.task.grid(row=3, column=2, sticky='w')

        self.batchnamelabel = Label(self.master, text='Batchname: ')
        self.batchnamelabel.grid(row=4, sticky='e')
        self.batchname = Entry(self.master, text='%s' % self.filename, bd=5, width=40, textvariable=self.filename, state=self.batchnamestat)
        self.batchname.grid(row=4, column=2)
        self.batchname.bind('<Button-1>', self.filedialog)

        self.RefIdlabel = Label(self.master, text='RefId: ')
        self.RefIdlabel.grid(row=5, sticky='e')
        self.RefId = Entry(self.master, text='%s' % self.RefIdvar, bd=5, textvariable=self.RefIdvar, state=self.RefIdstat)
        self.RefId.grid(row=5, column=2, sticky='w')

        self.ClusterIdlabel = Label(self.master, text='ClusterId: ')
        self.ClusterIdlabel.grid(row=6, sticky='e')
        self.ClusterId = Entry(self.master, text='%s' % self.ClusterIdvar, bd=5, textvariable=self.ClusterIdvar, state=self.ClusterIdstat)
        self.ClusterId.grid(row=6, column=2, sticky='w')

        self.timestartlabel = Label(self.master, text='Time Start: ')
        self.timestartlabel.grid(row=7, sticky='e')
        self.timestart = Entry(self.master, bd=5, text='%s' % self.ts, textvariable=self.ts, state=self.tsstat)
        self.timestart.grid(row=7, column=2, sticky='w')

        self.timeendlabel = Label(self.master, text='Time End: ')
        self.timeendlabel.grid(row=8, sticky='e')
        self.timeend = Entry(self.master, bd=5, text='%s' % self.te, textvariable=self.te, state=self.testat)
        self.timeend.grid(row=8, column=2, sticky='w')

        self.timespentlabel = Label(self.master, text='Time Spent: ')
        self.timespentlabel.grid(row=9, sticky='e')
        self.timespent = Entry(self.master, bd=5, text='%s' % self.tsp, textvariable=self.tsp, state=self.tspstat)
        self.timespent.grid(row=9, column=2, sticky='w')

        self.productivitylabel = Label(self.master, text='Productivity: ')
        self.productivitylabel.grid(row=10, sticky='e')
        self.productivity = Entry(self.master, bd=5, text='%s' % self.prod, textvariable=self.prod, state=self.prodstat, width=25)
        self.productivity.grid(row=10, column=2, sticky='w')
        self.productivity.bind('<Button-1>', self.filedialog)

        self.uom = OptionMenu(self.master,self.uomvar,*self.uomlist)
        self.uom.config(state=self.uomstat)
        self.uom.grid(row=10, column=2, sticky='e')
            
        self.statuslabel = Label(self.master, text='Status: ')
        self.statuslabel.grid(row=11,sticky='e')
        self.status = OptionMenu(self.master,self.statusvar,*self.statuslist)
        self.status.config(state=self.statusstat)
        self.status.grid(row=11, column=2, sticky='w')
            
        self.remarkslabel = Label(self.master, text='Remarks: ')
        self.remarkslabel.grid(row=12,sticky='e')
        self.remarks = Entry(self.master, bd=5, width=40, text='%s' % self.remarkstext, textvariable=self.remarkstext, state=self.remarksstat)
        self.remarks.grid(row=12, column=2)

        self.commands = {'Taskin':self.taskin, 'Taskout':self.taskout, 'OK':self.reset, 'Get':self.get_edit, 'Edit':self.edit, 'Update':self.update}
        self.OK = Button(self.master, text=self.OKtitle, command=self.commands[self.OKtitle], width=10)
        self.OK.grid(row=13, column=2, sticky='e', padx=10, pady=10)
        
        self.pts = Label(self.master, text='PTS-BDU, Innodata', fg='blue')
        self.pts.grid(row=13, sticky='w', padx=10, pady=10, columnspan=3)

        self.error=Label(self.master, text='%s!' % self.errormsg, fg=self.errorfg, bg=self.errorbg, wraplength=300)
        self.error.grid(row=14, column=0, columnspan=3, sticky='we', padx=10)
        self.Gui()

    def Gui(self):
        if self.checkongoing == 1:
            self.check_my_ongoing()
            self.checkongoing = 0
        return

    def get_xfiletypes(self):
        inifiletypes = {'redravel':'redravel', 'xml':'xml', '*':'all', 'tif':'tif', 'jpg':'jpg', 'manifest':'manifest'}
        x = 0
        if self.OK['text']=='Taskin':
            b = self.initask_split[3].split('; ')
       
        elif self.OK['text']=='Taskout':
            b = self.initask_split[4].split('; ')
        for a in b:
            if x == 0:
                self.xfiletypes = '%s files' % inifiletypes[a], '*.%s' % a
            else:
                self.xfiletypes = self.xfiletypes, ('%s files' % inifiletypes[a], '*.%s' % a)
            x += 1
        return

    def get_filename_init(self):
        self.tasknum=self.taskvar.get()[:self.taskvar.get().find(' ')]
        self.f = self.filename.get().split('" "')
        self.f[0] = self.f[0][1:]
        self.f[-1] = self.f[-1][:-1]
        return

    def get_productivity(self):
        p = 0
        missing_taskout_file = []
        for a in self.f:
            taskout_file = '%s.redravel' % a
            taskout_path = '%s/%s.redravel' % (self.path,a)
            if os.path.exists(taskout_path):
                name=os.path.split(taskout_path)[1]
                file_ext=os.path.splitext(name)[1]
                print('%s: %s' % ('file_ext', file_ext))
                if file_ext.upper()=='.REDRAVEL':
                    con_r = lite.connect(taskout_path)
                    with con_r:
                        cur_r = con_r.cursor()
                        cur_r.execute('SELECT count(pageno) from stats_pages_seen')
                        p=cur_r.fetchone()[0]
                        self.uomvar.set('%s %s' % (2,uomlist_dict[str(2)]))
                else:
                    p += os.stat(taskout_path).st_size
                    self.uomvar.set('%s %s' % (1,uomlist_dict[str(1)]))
            else:
                missing_taskout_file.append(taskout_file)
        if len(missing_taskout_file) > 0:
            self.error.configure(text = 'Missing taskout file: %s' % missing_taskout_file)
            self.error.configure(bg='red')
        self.prod.set(p)
        return

    def filedialog(self, event):
        if event.widget['text']=='PY_VAR0' and self.OK['text']=='Taskout':
            return
        self.check_task()
        open_dialog=0
        if self.is_task_OK==1:
            self.tasknum=self.taskvar.get()[:self.taskvar.get().find(' ')]
            self.initask_split=initasklist_dict[self.tasknum].split(', ')
            if self.OK['text']=='Taskin':
                self.path = self.initask_split[1]
                self.get_xfiletypes()
                open_dialog=1
            elif self.OK['text']=='Taskout':
                self.get_filename_init()
                self.path = '%s' % self.initask_split[2]
                self.get_xfiletypes()
                status=self.statusvar.get()[:self.statusvar.get().find(' ')]
                if status=='3':
                    self.error.configure(text = 'Fill-up status first')
                    self.error.configure(bg='red')
                elif status=='1':
                    self.path = '//mandaue.innodata.net/mdedfs01/RLIDG-P-EDIT/Working/RED_Markup_Pending'
                    open_dialog=1
                else:
                    open_dialog=1
        if open_dialog==1:
            print('%s: %s' % ('self.path', self.path))
            if self.OK['text']=='Taskout' and os.path.exists(self.path):
                self.get_productivity()
            else:
                self.get_xfiletypes()
                if self.tasknum=='6000':#deactivate parsing
                    chosen_file=filedialog.askopenfilenames(initialdir = self.path, title = "Select files", filetypes = self.xfiletypes)
                else:
                    chosen_file=(filedialog.askopenfilename(initialdir = self.path, title = "Select a file", filetypes = self.xfiletypes),)
                if len(chosen_file):
                    filenames=''
                    x = 0
                    for a in chosen_file:
                        if x==0:
                            filenames='"%s"' % os.path.splitext(os.path.split(a)[1])[0]
                        else:
                            filenames='"%s" %s' % (os.path.splitext(os.path.split(a)[1])[0], filenames)
                        x += 1
                    self.filename.set(filenames)
        return            
            
    def edit_init(self):
        self.error.configure(bg=self.defaultbg)
        self.timestart.configure(text='')
        self.timeend.configure(text='')
        self.timespent.configure(text='')
        self.productivity.configure(text='')
        self.uom.configure(text='')
        self.status.configure(text='')
        if self.editvar.get()==1:
            self.editvar.set(self.editvar.get())
            self.uservar.set(self.uservar.get())
            self.taskvar.set(self.taskvar.get())
            self.filename.set(self.filename.get())
            self.OK.configure(text='Get')
            self.OKtitle='Get'
            self.user.configure(stat=NORMAL)
            self.RefId.configure(stat=NORMAL)
            self.remarks.configure(stat=DISABLED)
            self.task.configure(stat=NORMAL)
            self.batchname.configure(stat=NORMAL)
        else:
            self.RefIdvar.set('')
            self.RefId.configure(stat=DISABLED)
            self.ClusterIdvar.set('')
            self.ClusterId.configure(stat=DISABLED)
            self.OK.configure(text='Taskin')
            self.check_my_ongoing()
        self.Gui
        return

    def get_edit(self):
        is_get_edit_OK = 0
        if self.RefIdvar.get() != '':
            self.RefIdvar.set(self.RefIdvar.get())
            is_get_edit_OK = 1
        elif self.uservar.get() == '':
            self.error.configure(text = 'User should not be empty if RefId is blank')
            self.error.configure(bg='red')
        elif self.taskvar.get() == '' or self.taskvar.get() =='NULL':
            self.error.configure(text = 'Task should not be empty or NULL if RefId is blank')
            self.errorbg.configure(bg='red')
        elif self.filename.get() == '':
            self.error.configure(text = 'Batch should not be empty if RefId is blank')
            self.errorbg.configure(bg='red')
        else:
            is_get_edit_OK = 1
        if is_get_edit_OK == 1:
            self.error.configure(bg=self.defaultbg)
            self.get_from_db('get_edit')
            if len(self.result)==1:
                self.uservar.set(self.result[0][2])
                t = self.result[0][7]
                self.taskvar.set('%s %s' % (t,initasklist_dict[str(t)].split(', ')[0]))
                self.filename.set(self.result[0][1])
                ts = self.result[0][3]
                self.ts.set(self.datetime_strptime_to_Guiformat(self.datetime_strptime(ts)))
                self.tsstat=NORMAL
                te = self.result[0][4]
                self.te.set(self.datetime_strptime_to_Guiformat(self.datetime_strptime(te)))
                self.testat=NORMAL
                self.tsp.set(self.result[0][5])
                self.tspstat=DISABLED
                self.prod.set(self.result[0][8])
                self.prodstat=NORMAL
                u = self.result[0][9]
                self.uomvar.set('%s %s' % (u,uomlist_dict[str(u)]))
                self.uomstat=NORMAL
                s = self.result[0][6]
                self.statusvar.set('%s %s' % (s,inistatuslist_dict[str(s)]))
                self.statusstat=NORMAL
                self.remarkstext.set(self.result[0][10])
                self.remarksstat=NORMAL
                self.RefIdvar.set(self.result[0][0])
                self.RefIdstat=DISABLED
                self.ClusterIdvar.set(self.result[0][11])
                self.ClusterIdstat=DISABLED
                self.OKtitle='Update'
            elif len(self.result)>1:
                print(self.result)
                self.error.configure(text='There are multiple record match, select RefId')
                self.error.configure(bg='red')
            elif len(self.result)==0:
                print(self.result)
                self.error.configure(text='No record match')
                self.error.configure(bg='red')
        self.Gui()
        return

    def update(self):
        ts = self.ts.get()
        te = self.te.get()
        is_ts_te_OK = 0
        if ts > te:
            self.error.configure(text='DateTimeStart %s should be less than DateTimeEnd %s' % (ts, te))
            self.error.configure(bg='red')
        else:
            is_ts_te_OK = 1
        is_register_OK = 0
        if is_ts_te_OK == 1:
            registry=self.get_from_db('get_registration')
            if len(registry)==0:
                self.error.configure(text='Batch %s is not registered' % (self.filename.get()))
                self.error.configure(bg='red')
            else:
                is_register_OK = 1
        if is_register_OK == 1:
            self.filename.set(self.filename.get())
            self.error.configure(bg=self.defaultbg)
            ts = self.Guiformat_to_dts(self.ts.get())
            self.start_time = self.datetime_strptime(ts)
            te = self.Guiformat_to_dts(self.te.get())
            self.end_time = self.datetime_strptime(te)
            self.tsp.set(self.get_timespent())
            self.update_db('update')
            self.editvar.set(0)
            self.reset()
        self.Gui()
        return

    #not used yet, still for reserch
    def validate(self, action, index, value_if_allowed, prior_value, text, validation_type, trigger_type, widget_name):
        if text in '0123456789,':
            try:
                float(value_if_allowed)
                return True
            except ValueError:
                return False
        else:
            return False

    def datetime_strptime(self, dt):
        return(datetime.strptime(str(dt), "%Y-%m-%d %H:%M:%S"))

    def datetime_strptime_to_Guiformat(self, dts):
        return(dts.strftime('%m/%d/%y %H:%M:%S'))

    def Guiformat_to_dts(self, gdt):
        return(datetime.strptime(str(gdt), "%m/%d/%y %H:%M:%S"))

    def check_my_ongoing(self):
        self.is_my_ongoing = self.get_from_db('my_ongoing')
        if len(self.is_my_ongoing) > 0:
            self.uomvar.set('NULL')
            filenames=''
            x = 0
            for a in self.result:
                if x==0:
                    filenames='"%s"' % a[0]
                else:
                    filenames='"%s" %s' % (a[0], filenames)
                x += 1
            self.filename.set(filenames)
            ts = self.datetime_strptime(self.result[0][1])
            self.ts.set(self.datetime_strptime_to_Guiformat(ts))
            self.start_time = time.strptime(str(self.ts.get()), "%m/%d/%y %H:%M:%S")
            self.start_time = time.mktime(self.start_time)
            t = self.result[0][2]
            self.taskvar.set('%s %s' % (t,initasklist_dict[str(t)].split(', ')[0]))
            self.statusvar.set('%s %s' % ('3',inistatuslist_dict['3']))
            self.remarkstext.set(self.result[0][3])
            self.RefIdvar.set(self.result[0][4])
            self.ClusterIdvar.set(self.result[0][5])
            self.OK.configure(text='Taskout')
            self.OK.configure(command=self.commands[self.OK['text']])
            self.batchname.configure(stat=DISABLED)
            self.task.configure(stat=DISABLED)
            self.status.configure(stat=NORMAL)
            self.uom.configure(stat=NORMAL)
            self.productivity.configure(stat=NORMAL)
        else:
            self.is_my_ongoing=0
        return

    def get_timestart(self):
        self.start_time=time.time()
        self.start_timestamp = datetime.fromtimestamp(self.start_time)
        self.datetimestart = self.start_timestamp.strftime('%m/%d/%y %H:%M:%S')
        return(self.datetimestart)
                     
    def get_timeend(self):
        self.end_time = time.time()
        self.end_timestamp = datetime.fromtimestamp(self.end_time)
        self.datetimeend = self.end_timestamp.strftime('%m/%d/%y %H:%M:%S')
        return(self.datetimeend)
    
    def get_timespent(self):
        elapsed = self.end_time - self.start_time
        #below is interim solution, need the root cause and fix
        try:
            d = datetime(1,1,1) + timedelta(seconds=int(elapsed))
        except:
            d = datetime(1,1,1) + elapsed
        elapsed = "%d:%d:%d:%d" % (d.day-1, d.hour, d.minute, d.second)
        return(elapsed)

    def check_task(self):
        self.is_task_OK=0
        if self.taskvar.get() not in self.tasklist:
            self.error.configure(text='Valid task is %s' % self.tasklist)
            self.error.configure(bg='red')
        else:
            self.taskvar.set(self.taskvar.get())
            self.is_task_OK=1
            self.error.configure(bg=self.defaultbg)
        return

    def taskin(self):
        is_filename_OK=0
        if self.filename.get()=='':
            self.error.configure(text='Batch is required')
            self.error.configure(bg='red')
        else:
            self.get_filename_init()
            for a in self.f:
                if not(re.match('^\d{14,14}$',a)):
                    self.error.configure(text='Batch %s does not conform to 14-digit barcode' % self.filename.get())
                    self.error.configure(bg='red')
                else:
                    is_filename_OK=1
        self.is_task_OK=0
        if is_filename_OK==1:
            self.check_task()
        is_ongoing_OK=0
        if self.is_task_OK==1:
            is_ongoing=self.get_from_db('check_if_ongoing')
            if len(is_ongoing)>0:
                self.error.configure(text='Batch %s is currently ongoing on "%s" task by %s' % (self.filename.get(),self.taskvar.get(),self.result[0][2]))
                print(self.result)
                self.error.configure(bg='red')
            else:
                is_ongoing_OK=1
        is_done_OK=0
        if is_ongoing_OK==1:
            is_done=self.get_from_db('check_if_done')
            if len(is_done)>0:
                self.error.configure(text='Batch %s is already done on "%s" task by %s' % (self.filename.get(),self.taskvar.get(),self.result[0][2]))
                self.error.configure(bg='red')
            else:
                is_done_OK=1
        if is_done_OK==1:
            self.get_filename_init()
            if len(self.f) > 1 and self.tasknum!='6000':#deactivate parsing
                self.error.configure(text='Only single file is allowed on task')
                self.error.configure(bg='red')
            elif self.ts.get()=='':
                print("Batchname: %s" % self.filename.get())
                self.ts.set(self.get_timestart())
                print("Taskin: %s" % self.ts.get())
                self.OK.configure(text='Taskout')
                self.OK.configure(command=self.commands[self.OK['text']])
                self.batchname.configure(stat = DISABLED)
                self.task.configure(stat=DISABLED)
                self.status.configure(stat=NORMAL)
                self.uom.configure(stat=NORMAL)
                self.remarks.configure(stat=NORMAL)
                self.remarkstext.set(self.remarkstext.get())
                self.check_remarkstext()
                self.error.configure(bg=self.defaultbg)
                self.productivity.configure(stat=NORMAL)
                self.statusvar.set('3 Ongoing')
                self.update_db('taskin')
        self.Gui()
        return

    def check_remarkstext(self):
        r = self.remarkstext.get()
        if r != '':
            self.remarkstext.set(self.remarkstext.get())
            print("Remarks: %s" % r)
        return

    def taskout(self):
        is_status_OK=0
        s = self.statusvar.get()
        self.statusvar.set(s)
        if s == '3 Ongoing':
            self.error.configure(text='Status "3 Ongoing" is invalid on taskout')
            self.error.configure(bg='red')
        else:
            is_status_OK=1
        is_prod_OK=0
        if is_status_OK==1:
            p = self.prod.get()
            self.prod.set(p)
            if p == '':
                self.error.configure(text='Productivity is required during taskout')
                self.error.configure(bg='red')
            elif p.isdigit() == False:
                self.error.configure(text='Productivity should be whole number')
                self.error.configure(bg='red')
                pass
            else:
                is_prod_OK=1
        is_uom_OK=0
        if is_prod_OK == 1:
            u = self.uomvar.get()
            self.uomvar.set(u)
            if u not in self.uomlist:
                self.error.configure(text='uom should be either %s' % self.uomlist)
                self.error.configure(bg='red')
            else:
                is_uom_OK=1
        if is_uom_OK==1:
            self.te.set(self.get_timeend())
            print("Taskout: %s" % self.te.get())
            self.tsp.set(self.get_timespent())
            print("Time Spent: %s" % self.tsp.get())
            self.OK.configure(text='OK')
            self.OK.configure(command=self.commands[self.OK['text']])
            self.check_remarkstext()
            self.remarks.configure(stat=DISABLED)
            self.uom.configure(stat=DISABLED)
            self.status.configure(stat=DISABLED)
            self.batchname.configure(stat=DISABLED)
            self.productivity.configure(stat=DISABLED)
            self.error.configure(bg=self.defaultbg)
            self.update_db('taskout')
        return

    def reset(self):
        self.batchname.configure(stat=NORMAL)
        self.remarks.configure(stat=NORMAL)
        self.task.configure(stat=NORMAL)
        self.filename.set('')
        self.ts.set('')
        self.te.set('')
        self.tsp.set('')
        self.prod.set('')
        self.OK.configure(text='Taskin')
        self.OK.configure(command=self.commands[self.OK['text']])
        self.error.configure(bg=self.defaultbg)
        self.remarkstext.set('')
        self.RefIdvar.set('')
        self.ClusterIdvar.set('')
        self.statusvar.set('NULL')
        self.taskvar.set('NULL')
        self.uomvar.set('NULL')
        self.user.configure(stat=DISABLED)
        self.timestart.configure(stat=DISABLED)
        self.timeend.configure(stat=DISABLED)
        self.status.configure(stat=DISABLED)
        self.uom.configure(stat=DISABLED)
        self.productivity.configure(stat=DISABLED)
        #self.Gui()
        return

    def connect_db(self):
        if dbsource == 'SQLserver':
            self.connection = pypyodbc.connect('Driver=SQL Server;' 'Server=%s;' 'Database=%s;' 'uid=%s;pwd=%s' % (sql_server, sql_db, sql_uid, sql_pwd))
        elif dbsource == 'SQLite':
            self.connection = lite.connect(sqlite_db)
        self.cursor = self.connection.cursor()
        return(self.connection)
       
    def get_from_db(self, curexecute):
        self.connect_db()
        self.get_filename_init()
        for file in self.f:
            if curexecute=='my_ongoing':
                self.cursor.execute('SELECT BatchName, DateTimeStarted, TaskId, Remarks, RefId, ClusterId FROM ' + self.TITOtable + ' WHERE UserId=? AND StatusId=3', (user,))
            elif curexecute=='get_registration':
                self.cursor.execute('SELECT ID, BatchName, "Process Type" from ' + self.shipmenttable + ' WHERE BatchName=?', (file,))
            elif curexecute=='check_if_ongoing':
                task=self.taskvar.get().split(' ')[0]
                self.cursor.execute('SELECT RefId, BatchName, UserId, DateTimeStarted, DateTimeEnded, Remarks, RefId from ' + self.TITOtable + ' WHERE StatusId=3 AND TaskId=? AND BatchName=?', (task,file))
            elif curexecute=='check_if_done':
                self.cursor.execute('SELECT RefId, BatchName, UserId, DateTimeStarted, DateTimeEnded, Remarks, RefId from ' + self.TITOtable + ' WHERE StatusId=2 AND TaskId=? AND BatchName=?', (self.taskvar.get().split(' ')[0],file))
            elif curexecute=='get_edit':
                if self.RefIdvar.get()!='':
                    self.cursor.execute('SELECT RefId, BatchName, UserId, DateTimeStarted, DateTimeEnded, TimeSpent, StatusId, TaskId, Productivity, uom, Remarks, ClusterId from ' + self.TITOtable + ' WHERE RefId=?', (self.RefIdvar.get(),))
                else:
                    self.cursor.execute('SELECT RefId, BatchName, UserId, DateTimeStarted, DateTimeEnded, TimeSpent, StatusId, TaskId, Productivity, uom, Remarks, RefId from ' + self.TITOtable + ' WHERE UserId=? AND TaskId=? AND BatchName=?', (self.uservar.get(),self.tasknum,file))
            elif curexecute=='is_edit_user':
                self.cursor.execute('SELECT id, UserId from ' + self.Usertable + ' WHERE UserId=?', (user,))
            self.result=self.cursor.fetchall()
        self.connection.close()
        return(self.result)

    def update_db(self, curexecute):
        self.connect_db()
        self.get_filename_init()
        is_cluster_off=0
        for file in self.f:
            if curexecute=='taskin':
                self.cursor.execute('SELECT MAX(RefId) FROM ' + self.TITOtable + '')
                RefId=self.cursor.fetchone()[0]
                if RefId is None:
                    RefId = 1
                else:
                    RefId += 1
                if self.tasknum=='6000':#deactivate multi-file in parsing
                    self.cursor.execute('SELECT MAX(ClusterId) FROM ' + self.TITOtable + '')
                    ClusterId=self.cursor.fetchone()[0]
                    if ClusterId is None:
                        ClusterId = 1
                        is_cluster_off=1
                    elif is_cluster_off==0:
                        ClusterId += 1
                        is_cluster_off=1
                    self.ClusterIdvar.set(ClusterId)
                    self.cursor.execute('INSERT INTO ' + self.TITOtable + ' (RefId, UserId, BatchName, DateTimeStarted, Remarks, TaskId, StatusId, ClusterId) VALUES (?, ?, ?, ?, ?, ?, 3, ?)', (RefId, user, file, self.ts.get(), self.remarkstext.get(), self.tasknum, ClusterId))
                else:
                    self.cursor.execute('INSERT INTO ' + self.TITOtable + ' (RefId, UserId, BatchName, DateTimeStarted, Remarks, TaskId, StatusId) VALUES (?, ?, ?, ?, ?, ?, 3)', (RefId, user, file, self.ts.get(), self.remarkstext.get(), self.tasknum))
                self.RefIdvar.set(RefId)
            elif curexecute=='taskout':
                if self.tasknum=='6000':#deactivate multi-file in parsing
                    taskout_path = '%s/%s.%s' % (self.path,file,user)
                    try:
                        prod = os.stat(taskout_path).st_size
                    except:
                        error='Output file %s was not found' % taskout_path
                        self.error.configure(text='%s' % error)
                        self.error.configure(bg='red')
                        print(error)
                        prod = 0
                    if is_cluster_off == 0:
                        te = self.te.get() 
                        tsp = self.tsp.get()
                        is_cluster_off = 1
                    else:
                        te = self.ts.get()
                        tsp = '0:0:0:0'
                    self.cursor.execute('UPDATE ' + self.TITOtable + ' SET DateTimeEnded=?, TimeSpent=?, Productivity=?, uom=?, StatusId=?, Remarks=? WHERE BatchName=? AND UserId=? AND TaskId=? AND ClusterId=?', (te, tsp, prod, self.uomvar.get().split(' ')[0], self.statusvar.get().split(' ')[0], self.remarkstext.get(), file, self.uservar.get(), self.tasknum, self.ClusterId.get()))
                else:
                    self.cursor.execute('UPDATE ' + self.TITOtable + ' SET DateTimeEnded=?, TimeSpent=?, Productivity=?, uom=?, StatusId=?, Remarks=? WHERE BatchName=? AND UserId=? AND TaskId=? AND RefId=?', (self.te.get(), self.tsp.get(), self.prod.get(), self.uomvar.get().split(' ')[0], self.statusvar.get().split(' ')[0], self.remarkstext.get(), file, self.uservar.get(), self.tasknum, self.RefIdvar.get()))
            elif curexecute=='update':
                self.cursor.execute('UPDATE ' + self.TITOtable + ' SET UserId=?, TaskId=?, BatchName=?, DateTimeStarted=?, DateTimeEnded=?, TimeSpent=?, Productivity=?, uom=?, StatusId=?, Remarks=? WHERE RefId=?', (self.uservar.get(), self.tasknum, file, self.ts.get(), self.te.get(), self.tsp.get(), self.prod.get(), self.uomvar.get().split(' ')[0], self.statusvar.get().split(' ')[0], self.remarkstext.get(), self.RefIdvar.get()))
            if dbsource=='SQLServer':
                self.cursor.commit()
            else:
                self.connection.commit()
        self.connection.close()
        return
            
def process():
    root = Tk()
    frame = TITOGui(root)
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
        global titoable, inistatuslist, inistatuslist_dict, initasklist, initasklist_dict, scalc, dbsource, sql_server, sql_db, sql_uid, sql_pwd, sqlite_db, uomlist, uomlist_dict, project
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
        project = inivariables['project']
        sqlite_db = inivariables['sqlite_db']
        scalc = inivariables['scalc']
        scalcsplit = scalc.split('; ')
        for a in scalcsplit:
            if os.path.exists(a):
                scalc=a
                break
        cur.execute('SELECT id, status from "Status"')
        temp = cur.fetchall()
        inistatuslist = []
        for a,b in temp:
            inistatuslist.append('%s %s' % (a,b))
        inistatuslist_dict = {}
        for a in iter(temp):
            inistatuslist_dict[str(a[0])]=a[1]
        cur.execute('SELECT id, task, inputpath, outputpath, inputfiletype, outputfiletype from "Task"')
        tasktable = cur.fetchall()
        initasklist = []
        for a,b,c,d,e,f in tasktable:
            initasklist.append('%s %s' % (a,b))
        initasklist_dict = {}
        for a in iter(tasktable):
            initasklist_dict[str(a[0])]='%s, %s, %s, %s, %s' % (a[1], a[2], a[3], a[4], a[5])
        cur.execute('SELECT id, uom from "uom"')
        temp = cur.fetchall()
        uomlist = []
        for a,b in temp:
            uomlist.append('%s %s' % (a,b))
        uomlist_dict = {}
        for a in iter(temp):
            uomlist_dict[str(a[0])]=a[1]
    header()
    print(sql_server)
    process()

if __name__ == '__main__':
    main()
