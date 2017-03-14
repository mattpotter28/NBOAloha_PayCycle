# Matt Potter
# Created January 5 2017
# Last Edited February 7 2017
# NBOAloha.py

import configparser
import base64
import pypyodbc
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
from datetime import datetime

class MainApplication: # main window
    def __init__(self, master, note, conn, serv):
        # region constructor

        # SQL Connection
        MainApplication.connection = conn
        MainApplication.cursor = self.connection.cursor()
        self.SQLCommand = ("")
        self.serverName = serv

        # window and tabs creation
        self.master = master
        self.frame = ttk.Frame(self.master)

        # sets titles
        self.master.wm_title("Pay Cycle Setup")
        self.serverLab = tk.Label(self.frame, text="Connected to: " + self.serverName)
        self.serverLab.grid(column=0, row=0, padx=5, pady=5, sticky='W')
        self.tableButton = tk.Button(self.frame, text="View Consolidated Table", command=self.newTableWindow)
        self.tableButton.grid(column=1, row=0, padx=(125,5), pady=5, sticky='E')

        # tab creation
        aloha(self.master, note)
        nbo(self.master, note)

        # places completed frame and tabs on the grid
        self.frame.grid(column=0, row=0)
        note.grid(column=0, row=1, columnspan=4, padx=5, pady=5)

        #endregion constructor

    def newTableWindow(self):
        # region opens Table Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = cTableWindow(self.newWindow)
        # endregion opens Table Window

class aloha:
    def __init__(self, master, note):
        self.master = master
        self.aTab = ttk.Frame(note)
        note.add(self.aTab, text="Aloha", compound='top')
        self.fields = ('Location', 'ADP Store Code', 'Pay Cycle', 'Begin Date', 'End Date')
        self.createWidgets()

    def createWidgets(self):
        self.days = list(range(1,32))
        self.months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        self.monthsDict = {'Jan':'01', 'Feb':'02', 'Mar':'03', 'Apr':'04', 'May':'05', 'Jun':'06', 'Jul':'07', 'Aug':'08',
                       'Sep':'09', 'Oct':'10', 'Nov':'11', 'Dec':'12'}
        self.years = list(range(2017, 2080))
        for field in self.fields:
            if field == 'Location':
                self.LocationVariable = tk.StringVar()
                self.lLab = tk.Label(self.aTab, text="Location: ")
                self.lLab.grid(column=1, row=1, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select [strLocationReportHeading] FROM dbo.tbl_UnitDescription")
                MainApplication.cursor.execute(self.SQLCommand)
                self.siteNames = MainApplication.cursor.fetchall()
                self.siteNames = [str(s).strip('{(\"\',)}') for s in self.siteNames]
                self.siteNames.sort()
                self.lEntry = ttk.Combobox(self.aTab, textvariable=self.LocationVariable,
                                                      value=self.siteNames, state="readonly")
                self.lEntry.grid(column=2, columnspan=4, row=1, padx=(44,5), pady=5, sticky="ew")
                self.lEntry.bind('<<ComboboxSelected>>', self.locationSelect)

            elif field == 'ADP Store Code':
                # region ADP Code Field
                self.sLab = tk.Label(self.aTab, text="ADP Store Code: ")
                self.sLab.grid(column=1, row=2, padx=5, pady=5, sticky='W')
                self.ADPStoreCodeVariable = tk.StringVar()
                self.sEntry = tk.Entry(self.aTab, textvariable=self.ADPStoreCodeVariable, width=30)
                self.sEntry.grid(column=2, columnspan=4, row=2, padx=44, pady=5, sticky='W')
                # endregion ADP Code Field

            elif field == 'Pay Cycle':
                # region Pay Cycle Field
                self.cLab = tk.Label(self.aTab, text="Pay Cycle: ")
                self.cLab.grid(column=1, row=3, padx=5, pady=5, sticky='W')
                self.PayCycleVariable = tk.StringVar()
                self.cEntry = ttk.Combobox(self.aTab, textvariable=self.PayCycleVariable,
                                           values=(1,2,3,4), state="readonly")
                self.cEntry.grid(column=2, columnspan=4, row=3, padx=(44,5), pady=5, sticky='W')
                # endregion Pay Cycle Field

            elif field == 'Begin Date':
                # region Begin Date Field
                self.bLab = tk.Label(self.aTab, text="Begin Date: ")
                self.bLab.grid(column=1, row=4, padx=5, pady=5, sticky='W')
                self.bmonthVariable = tk.StringVar()
                self.bdayVariable = tk.StringVar()
                self.byearVariable = tk.StringVar()
                self.bmonthEntry = ttk.Combobox(self.aTab, textvariable=self.bmonthVariable, values=self.months, width=4)
                self.bmonthEntry.grid(column=2, row=4, padx=(44,5), pady=5, sticky='W')
                self.bdayEntry = ttk.Combobox(self.aTab, textvariable=self.bdayVariable, values=self.days, width=3)
                self.bdayEntry.grid(column=3, row=4, padx=5, pady=5, sticky='W')
                self.byearEntry = ttk.Combobox(self.aTab, textvariable=self.byearVariable, values=self.years, width=5)
                self.byearEntry.grid(column=4, row=4, padx=5, pady=5, sticky='W')
                # endregion Begin Date Field


            elif field == 'End Date':
                # region Begin Date Field
                self.eLab = tk.Label(self.aTab, text="End Date: ")
                self.eLab.grid(column=1, row=5, padx=5, pady=5, sticky='W')
                self.emonthVariable = tk.StringVar()
                self.edayVariable = tk.StringVar()
                self.eyearVariable = tk.StringVar()
                self.emonthEntry = ttk.Combobox(self.aTab, textvariable=self.emonthVariable, values=self.months, width=4)
                self.emonthEntry.grid(column=2, row=5, padx=(44,5), pady=5, sticky='W')
                self.edayEntry = ttk.Combobox(self.aTab, textvariable=self.edayVariable, values=self.days, width=3)
                self.edayEntry.grid(column=3, row=5, padx=5, pady=5, sticky='W')
                self.eyearEntry = ttk.Combobox(self.aTab, textvariable=self.eyearVariable, values=self.years, width=5)
                self.eyearEntry.grid(column=4, row=5, padx=5, pady=5, sticky='W')
                # endregion Begin Date Field

        # endregion Label/Option Menu creation

        self.submitButton = tk.Button(self.aTab, text="Submit", command=self.submit)
        self.submitButton.grid(column=6, columnspan=1, row=6, padx=5, pady=5)
        self.tableButton = tk.Button(self.aTab, text="View Table", command=self.newTableWindow)
        self.tableButton.grid(column=5, columnspan=1, row=6, padx=5, pady=5)
        self.cancelButton = tk.Button(self.aTab, text="Close", command=self.master.destroy)
        self.cancelButton.grid(column=1, columnspan=2, row=6, padx=5, pady=5, sticky='W')

    def locationSelect(self, oth):
        # check current location selection
        location = self.siteNames[self.lEntry.current()]

        # clear current values from entries
        self.sEntry.delete(0, 'end')
        self.cEntry.set('')
        self.bmonthEntry.set('')
        self.bdayEntry.set('')
        self.byearEntry.set('')
        self.emonthEntry.set('')
        self.edayEntry.set('')
        self.eyearEntry.set('')

        # find corresponding pay cycle and ADP code to name
        self.SQLCommand = (
        "select [strADP_Code] from [POSLabor].[dbo].[tbl_UnitDescription] where [strLocationReportHeading] = \'" + location + "\'")
        MainApplication.cursor.execute(self.SQLCommand)
        adp = MainApplication.cursor.fetchall()
        self.sEntry.insert(0, adp)

        self.SQLCommand = (
        "select [intPayCycle] from [POSLabor].[dbo].[tbl_UnitDescription] where [strLocationReportHeading] = \'" + location + "\'")
        MainApplication.cursor.execute(self.SQLCommand)
        cycleNum = str(MainApplication.cursor.fetchall()).strip("[((,))]")
        self.cEntry.set(cycleNum)

    def checkDates(self, bm, bd, by, em, ed, ey):
        max = datetime(month=6, day=7, year=2079) # maximum date of June 6, 2079
        try:
            self.bdate = datetime(month=int(bm), day=int(bd), year=int(by))
            if self.bdate > max:
                raise
        except:
            self.mBox = tk.messagebox.showinfo("Error!", "Invalid Begin Date")
        try:
            self.edate = datetime(month=int(em), day=int(ed), year=int(ey))
            if self.edate > max:
                raise
            if self.edate < self.bdate:
                raise
        except:
            self.mBox = tk.messagebox.showinfo("Error!", "Invalid End Date")

    def submit(self):
        ## Location conversion ##
        self.loc = self.LocationVariable.get()
        self.SQLCommand = ("SELECT [iLocationID] FROM [POSLabor].[dbo].[tbl_UnitDescription] where [strLocationReportHeading] like '" + self.loc + "';")
        MainApplication.cursor.execute(self.SQLCommand)
        self.loc = MainApplication.cursor.fetchone()
        self.loc = str(self.loc).strip("(,)")

        ## ADP
        self.adp = self.ADPStoreCodeVariable.get()

        ## Pay Cycle
        self.payc = self.PayCycleVariable.get()

        ## Begin Date
        self.bmonth = self.monthsDict[self.bmonthVariable.get()]
        self.bday = "%02d" % int(self.bdayVariable.get())
        self.byear = self.byearVariable.get()
        self.bdate = "{}-{}-{} 00:00:00".format(self.byear, self.month, self.day)

        ## End Date
        self.emonth = self.monthsDict[self.emonthVariable.get()]
        self.eday = "%02d" % int(self.edayVariable.get())
        self.eyear = self.eyearVariable.get()
        # self.edate = "{}-{}-{} 00:00:00".format(self.eyear, self.month, self.day)

        ## check if dates are valid
        self.checkDates(self.bmonth, self.bday, self.byear, self.emonth, self.eday, self.eyear)

        self.insertSQL(self.loc, self.adp, self.payc, self.bdate, self.edate)

    def insertSQL(self, loc, adp, payc, bdate, edate):
        try:
            # region Statement submittion
            self.SQLCommand = ("INSERT INTO POSLabor.Labor.LBR_StorePayCycle "
                               "VALUES ("+loc+", '"+adp+"', "+payc+", '"+bdate+"', '"+edate+"');")
            print(self.SQLCommand)
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()
            self.mBox = tk.messagebox.showinfo("Success!","Import Complete")
            # endregion Statement submittion

        except:
            # region submition error
            self.mBox = tk.messagebox.showinfo("Error!","Import Failed")
            print("Failed Command: " + self.SQLCommand)
            MainApplication.connection.rollback()
            # endregion submition error

        self.master.destroy()

    def newTableWindow(self):
        # region opens Table Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = alohaTableWindow(self.newWindow)
        # endregion opens Table Window

class alohaTableWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("View Table")
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        self.table = ttk.Treeview(self.frame)
        self.table.grid(column=1, row=1, columnspan=4, padx=(5,0), pady=5)
        self.table['columns'] = ('adp', 'payCycle', 'bDate', 'eDate')
        self.table.heading('#0', text="Location")
        self.table.column('#0', width=75)
        self.table.heading('adp', text="ADP Store Code")
        self.table.column('adp', width=100)
        self.table.heading('payCycle', text="Pay Cycle")
        self.table.column('payCycle', width=100)
        self.table.heading('bDate', text="Begin Date")
        self.table.column('bDate', width=125)
        self.table.heading('eDate', text="End Date")
        self.table.column('eDate', width=125)

        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.table.yview)
        self.table.configure(yscroll=sb.set)
        sb.grid(row=1, column=5, sticky="NS", padx=(0,5), pady=5)

        self.SQLCommand = ("SELECT * FROM POSLabor.Labor.LBR_StorePayCycle")
        MainApplication.cursor.execute(self.SQLCommand)
        for row in MainApplication.cursor:
            self.table.insert('', 'end', text=str(row[0]), values=(row[1], row[2], row[3], row[4]))

        self.exportButton = tk.Button(self.frame, text="Export...", command=self.export)
        self.exportButton.grid(column=2, row=2, padx=5, pady=5, sticky='E')
        self.closeButton = tk.Button(self.frame, text="Close", command=self.close_windows)
        self.closeButton.grid(column=3, row=2, padx=5, pady=5, sticky='W')

    def export(self):
        # prepare text file
        ftypes = [('Text file', '.txt'), ('All files', '*')]
        name = filedialog.asksaveasfilename(filetypes=ftypes, defaultextension=".txt")
        outFile = open(name, 'w')
        # prepare sql query
        self.SQLCommand = ("SELECT * FROM POSLabor.Labor.LBR_StorePayCycle")
        MainApplication.cursor.execute(self.SQLCommand)
        # fill text file
        for row in MainApplication.cursor:
            outFile.write(str(row[0])+"\t"+row[1]+"\t"+str(row[2])+"\t"+str(row[3])+"\t"+str(row[4])+"\n")
        outFile.close()

    def close_windows(self):
        self.master.destroy()

class nbo:
    def __init__(self, master, note):
        self.master = master
        self.frame = ttk.Frame(note)
        note.add(self.frame, text="NBO")
        self.fields = ('Location', 'Pay Group', 'Tip Share', 'Pay Cycle', 'ADP Store Code')
        self.createWidgets()

    def createWidgets(self):
        # region Label/Option Menu creation

        for field in self.fields:
            if field == 'Location':
                # region Location Field
                self.LocationVariable = tk.StringVar()
                self.lLab = tk.Label(self.frame, text="Location: ")
                self.lLab.grid(column=1, columnspan=2, row=1, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select [SiteName] FROM dbo.NBO_Sites")
                MainApplication.cursor.execute(self.SQLCommand)
                self.siteNames = MainApplication.cursor.fetchall()
                self.siteNames = [str(s).strip('{(\"\',)}') for s in self.siteNames]
                self.siteNames.sort()
                MainApplication.lEntry = ttk.Combobox(self.frame, textvariable=self.LocationVariable, value = self.siteNames, state="readonly")
                MainApplication.lEntry.grid(column=4, columnspan=6, row=1, padx=5, pady=5, sticky="ew")
                MainApplication.lEntry.bind('<<ComboboxSelected>>', self.locationSelect)
                # endregion Location Field

            elif field == 'Pay Group':
                # region Pay Group Field
                nbo.PayGroupVariable = tk.StringVar()
                self.gLab = tk.Label(self.frame, text="Pay Group: ")
                self.gLab.grid(column=1, columnspan=2, row=2, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select distinct [PayrollGroupName] from [POSLabor].[dbo].[NBO_PayGroup]")
                MainApplication.cursor.execute(self.SQLCommand)
                nbo.payGroups = MainApplication.cursor.fetchall()
                nbo.payGroups = [str(s).strip('{(\"\',)}') for s in nbo.payGroups]
                nbo.gEntry = ttk.Combobox(self.frame, textvariable=nbo.PayGroupVariable, values=self.payGroups, state="readonly")
                nbo.gEntry.grid(column=4, columnspan=6, row=2, padx=5, pady=5, sticky="ew")
                # endregion Pay Group Field

            elif field == 'Tip Share':
                # region Tip Share Field
                self.tLab = tk.Label(self.frame, text="Tip Share: ")
                self.tLab.grid(column=1, columnspan=2, row=3, padx=5, pady=5, sticky='W')
                self.TipShareVariable = tk.StringVar()
                self.tEntry = ttk.Combobox(self.frame, textvariable=self.TipShareVariable, values=("No Tip Share","Landrys Tip Share","NBO Tip Share"))
                self.tEntry.grid(column=4, columnspan=6, row=3, padx=5, pady=5, sticky='W')
                # endregion Tip Share Field

            elif field == 'Pay Cycle':
                # region Pay Cycle Field
                self.cLab = tk.Label(self.frame, text="Pay Cycle: ")
                self.cLab.grid(column=1, columnspan=2, row=4, padx=5, pady=5, sticky='W')
                self.PayCycleVariable = tk.StringVar()
                self.SQLCommand = ("SELECT LTRIM(RTRIM(RIGHT(BusinessCalendarName, 2 ))) as CycleNumber FROM [NBO_TRAIN].[dbo].[BusinessCalendars] WHERE BusinessCalendarName LIKE 'Payroll Cycle%'")
                MainApplication.cursor.execute(self.SQLCommand)
                MainApplication.payCycles = MainApplication.cursor.fetchall()
                self.cEntry = ttk.Combobox(self.frame, textvariable=self.PayCycleVariable, values=MainApplication.payCycles, state="readonly")
                self.cEntry.grid(column=4, columnspan=6, row=4, padx=5, pady=5, sticky='W')
                # endregion Pay Cycle Field

            elif field == 'ADP Store Code':
                # region ADP Code Field
                self.sLab = tk.Label(self.frame, text="ADP Store Code: ")
                self.sLab.grid(column=1, columnspan=2, row=5, padx=5, pady=5, sticky='W')
                self.ADPStoreCodeVariable = tk.StringVar()
                self.sEntry = tk.Entry(self.frame, textvariable=self.ADPStoreCodeVariable, width=30)
                self.sEntry.grid(column=4, columnspan=6, row=5, padx=5, pady=5, sticky='W')
                # endregion ADP Code Field

        # endregion Label/Option Menu creation

        # region Button creation
        self.tableButton = tk.Button(self.frame, text="View Table", command=self.newTableWindow)
        self.tableButton.grid(column=3, columnspan=2, row=6, padx=5, pady=5)
        self.addButton = tk.Button(self.frame, text="Add Pay Group", command=self.newAddWindow)
        self.addButton.grid(column=5, columnspan=2, row=6, padx=5, pady=5)
        self.editButton = tk.Button(self.frame, text="Edit Pay Group", command=self.newEditWindow)
        self.editButton.grid(column=7, columnspan=2, row=6, padx=5, pady=5)
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=9, columnspan=2, row=6, padx=5, pady=5)
        self.cancelButton = tk.Button(self.frame, text="Close", command=self.master.destroy)
        self.cancelButton.grid(column=1, columnspan=2, row=6, padx=5, pady=5, sticky='W')
        # endregion Button creation

    def locationSelect(self, oth):
        # check current location selection
        location = self.siteNames[MainApplication.lEntry.current()]
        location = str(location).replace("'", "%")

        # clear current values from entries
        self.gEntry.set('')
        self.tEntry.set('')
        self.cEntry.set('')
        self.sEntry.delete(0,'end')

        # find corresponding site number to name
        self.SQLCommand = ("select [SiteNumber] from dbo.NBO_Sites where [SiteName] like \'%"+location+"%\'")
        MainApplication.cursor.execute(self.SQLCommand)
        siteNumber = MainApplication.cursor.fetchall()
        siteNumber = str(siteNumber).strip("[(\',)]")

        # take in values based on site number
        self.SQLCommand = ("SELECT * FROM [POSLabor].[dbo].[NBO_PayCycleSetup] where [SiteNumber] like \'%"+siteNumber+"%\'")
        MainApplication.cursor.execute(self.SQLCommand)

        # fill entries with values
        for row in MainApplication.cursor:
            # 1 - PayGroupID: translate ID to name and match with combobox option
            self.SQLCommand = ("SELECT [PayrollGroupName] FROM [POSLabor].[dbo].[NBO_PayGroup] where [PayGroupID] = \'"+str(row[1])+"\'")
            MainApplication.cursor.execute(self.SQLCommand)
            groupName = str(MainApplication.cursor.fetchall()).strip("[(\',)]")
            self.gEntry.set(groupName)

            # 2 - TipShareLocation: translate to boolean to apply to checkbutton
            if int(row[2]) == 0:
                self.tEntry.set("No Tip Share")
            elif int(row[2]) == 1:
                self.tEntry.set("Landrys Tip Share")
            elif int(row[2]) == 2:
                self.tEntry.set("NBO Tip Share")

            # 3 - ADPStoreCode: fill entry with string
            self.sEntry.insert(0, str(row[3]))

            # 5 - NBOCalendarNumber: translate to paycycle and match with combobox option
            if row[5] == 1693843:
                self.cEntry.set('1')
            elif row[5] == 1693844:
                self.cEntry.set('2')
            elif row[5] == 1693845:
                self.cEntry.set('3')
            elif row[5] == 1837816:
                self.cEntry.set('4')
            elif row[5] == 1908266:
                self.cEntry.set('5')

    def newTableWindow(self):
        # region opens Table Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = nboTableWindow(self.newWindow)
        # endregion opens Table Window

    def newAddWindow(self):
        # region opens Add Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = addWindow(self.newWindow)
        # endregion opens Add Window

    def newEditWindow(self):
        # region opens Edit Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = editWindow(self.newWindow)
        # endregion region opens Edit Window

    def submit(self):
        # region SQLStatement creation

        ## Location conversion ##
        self.loc = self.LocationVariable.get()
        self.loc = self.loc.strip("(\",)")
        self.loc = str(self.loc).replace("'","%")
        self.SQLCommand = ("SELECT [SiteNumber] FROM [POSLabor].[dbo].[NBO_Sites] where [SiteName] like '" + self.loc + "';")
        MainApplication.cursor.execute(self.SQLCommand)
        self.loc = MainApplication.cursor.fetchone()
        self.loc = str(self.loc).strip("(,)")

        ## Pay Group conversion ##
        self.payg = self.PayGroupVariable.get()
        self.payg = self.payg.strip("(',)")
        self.SQLCommand = ("SELECT [PayGroupID] FROM [POSLabor].[dbo].[NBO_PayGroup] where [PayrollGroupName] like '" + self.payg + "';")
        MainApplication.cursor.execute(self.SQLCommand)
        self.payg = MainApplication.cursor.fetchone()
        self.payg = str(self.payg).strip("(,)")

        if(self.TipShareVariable.get() == "No Tip Share"):
            self.tip = 0
        elif(self.TipShareVariable.get() == "Landrys Tip Share"):
            self.tip = 1
        elif(self.TipShareVariable.get() == "NBO Tip Share"):
            self.tip = 2

        self.payc = self.PayCycleVariable.get()
        self.adp = self.ADPStoreCodeVariable.get()
        self.insertSQL(self.loc, self.payg, self.tip, self.payc, self.adp)

        # endregion SQLStatement creation

    def insertSQL(self, loc, payg, tip, payc, adp):
        try:
            # region Statement submittion
            self.SQLCommand = ("DECLARE @RT INT "\
                           "EXECUTE @RT = dbo.pr_NBO_PayCycleSetup_ADD " + self.loc + ", " + self.payg + ", " + str(self.tip) + ", '" + self.adp + "', 0, " + self.payc + " PRINT @RT") # command to add data
            print(self.SQLCommand)
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()
            self.mBox = tk.messagebox.showinfo("Success!","Import Complete")
            # endregion Statement submittion

        except:
            # region submition error
            self.mBox = tk.messagebox.showinfo("Error!","Import Failed")
            print("Failed Command: " + self.SQLCommand)
            MainApplication.connection.rollback()
            # endregion submition error

        self.master.destroy()

class nboTableWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("View Table")
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        self.table = ttk.Treeview(self.frame)
        self.table.grid(column=1, columnspan=6, row=1, rowspan=5, padx=(5,0), pady=5)
        self.table['columns'] = ('pgID', 'tipShare', 'adp', 'active', 'nbo')
        self.table.heading('#0', text="Site Name")
        self.table.column('#0', width=275)
        self.table.heading('pgID', text="Payroll Group Name")
        self.table.column('pgID', width=150)
        self.table.heading('tipShare', text="Tip Share")
        self.table.column('tipShare', width=100)
        self.table.heading('adp', text="ADP Store Code")
        self.table.column('adp', width=100)
        self.table.heading('active', text="Active Location")
        self.table.column('active', width=100)
        self.table.heading('nbo', text="Pay Cycle")
        self.table.column('nbo', width=100)

        self.SQLCommand = ("EXECUTE dbo.pr_NBO_PayCycleReport")
        MainApplication.cursor.execute(self.SQLCommand)
        for row in MainApplication.cursor:
            self.table.insert('', 'end', text=str(row[0]), values=(row[1], row[2], row[3], row[4], row[5]))
        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.table.yview)
        self.table.configure(yscroll=sb.set)
        sb.grid(row=1, column=9, rowspan=5, sticky="NS", pady=5)
        self.exportButton = tk.Button(self.frame, text="Export...", command=self.export)
        self.exportButton.grid(column=1, columnspan=3, row=6, padx=5, pady=5, sticky='E')
        self.closeButton = tk.Button(self.frame, text="Close", command=self.close_windows)
        self.closeButton.grid(column=4, columnspan=3, row=6, padx=5, pady=5, sticky='W')

    def export(self):
        # prepare text file
        ftypes = [('Text file', '.txt'), ('All files', '*')]
        name = filedialog.asksaveasfilename(filetypes=ftypes, defaultextension=".txt")
        outFile = open(name, 'w')
        # prepare sql query
        self.SQLCommand = ("EXECUTE dbo.pr_NBO_PayCycleReport")
        MainApplication.cursor.execute(self.SQLCommand)
        # fill text file
        for row in MainApplication.cursor:
            outFile.write(row[0]+"\t"+row[1]+"\t"+str(row[2])+"\t"+row[3]+"\t"+str(row[4])+"\t"+str(row[5])+"\n")
        outFile.close()

    def close_windows(self):
        self.master.destroy()

class addWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Add Pay Group")
        self.master.geometry("%dx%d%+d%+d" % (250, 100, 250, 125))
        self.newPayGroupName = tk.StringVar()
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        # region Label/Entry creation
        self.lab = tk.Label(self.frame, text="New Pay Group Name")
        self.lab.grid(column=1, columnspan = 4, row=1, padx=5, pady=5, sticky='W')
        self.entry = tk.Entry(self.frame, textvariable=self.newPayGroupName, width = 39)
        self.entry.grid(column=1, columnspan=4, row=2, padx=5, pady=5, sticky='W')
        # endregion Label/Entry creation

        # region Button creation
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=1, row=3, padx=5, pady=5, sticky='W')
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=4, row=3, padx=5, pady=5, sticky='E')
        # endregion Button creation

    def submit(self):
        # region name validation
        self.flag = False
        for name in nbo.payGroups:
            print(name+"\n"+self.newPayGroupName.get())
            if str(self.newPayGroupName.get()) == name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True
        # endregion name validation

        # region SQL submittion
        if self.flag == False:
            nbo.payGroups.insert((len(MainApplication.payGroups) + 1), self.newPayGroupName.get())
            self.SQLCommand = ("INSERT INTO [POSLabor].[dbo].[NBO_PayGroup] (PayrollGroupName, PayGroupID) " \
                          "VALUES ('" + str(self.newPayGroupName.get()) + "', " + str(len(nbo.payGroups)) + " );")
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

            # region resetting menus
            MainApplication.payGroups.insert(0,self.newPayGroupName)
            MainApplication.gEntry['values'] = MainApplication.payGroups
            MainApplication.PayGroupVariable.set(self.newPayGroupName.get())
            # endregion resetting menus

        # endregion SQL submittion

        self.master.destroy()

    def close_windows(self):
        self.master.destroy()

class editWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Edit Pay Group")
        self.master.geometry("%dx%d%+d%+d" % (411, 205, 250, 125))
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        # region Label/Entry creation

        self.lab1 = tk.Label(self.frame, text="Pay Group to Edit:")
        self.lab1.grid(column=1, columnspan=5, row=1, sticky='W', padx=5, pady=5)

        ## nameList
        self.nameList = tk.Listbox(self.frame, width=25)
        self.count = 0
        for name in nbo.payGroups:
            name = str(name).strip("(,)")
            self.nameList.insert(self.count, name)
            self.count = self.count + 1
        self.nameList.grid(row=2, rowspan=6, column=1, padx=(5,0), pady=5)
        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.nameList.yview)
        self.nameList.configure(yscroll=sb.set)
        sb.grid(row=2, column=2, rowspan=6, sticky="NS", pady=5)
        self.lab2 = tk.Label(self.frame, text="New Pay Group Name: ", anchor='w', width=30)
        self.lab2.grid(row=3, column=3, columnspan=4, padx=10, pady=5, sticky='S')
        self.NewPayGroupName = tk.StringVar()
        self.entry = tk.Entry(self.frame, textvariable=self.NewPayGroupName, width=30)
        self.entry.grid(row=4, column=3, columnspan=4, padx=10, sticky='NW')
        self.entry.grid_columnconfigure(3, weight=1)
        self.submitButton = tk.Button(self.frame, text="Submit Edit", command=lambda: self.submit(self.nameList.get(tk.ACTIVE), self.NewPayGroupName))
        self.submitButton.grid(column=6, columnspan=2, row=7, padx=5, pady=5, sticky='SE')
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=4, columnspan=2, row=7, padx=(5,0), pady=5, sticky='SE')
        # endregion Label/Entry creation

    def submit(self, OldPayGroupName, NewPayGroupName):
        # region name validation
        self.flag = False
        for name in nbo.payGroups:
            if str(NewPayGroupName.get()) in name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True
        # endregion name validation

        # region SQL submittion
        if self.flag == False:
            # use sql statement to UPDATE values to new values
            OldPayGroupName = str(OldPayGroupName).replace("'", "%")
            self.SQLCommand = ("UPDATE [POSLabor].[dbo].[NBO_PayGroup] " \
                          "SET [PayrollGroupName]='" + str(NewPayGroupName.get()) + "' " + \
                          "WHERE [PayrollGroupName] like '" + OldPayGroupName + "';")
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

            # region resetting menus
            OldPayGroupName = OldPayGroupName.strip('%')
            for index, item in enumerate(nbo.payGroups):
                if (OldPayGroupName in item):
                    nbo.payGroups[index] = str(NewPayGroupName.get())
                    nbo.gEntry.set(nbo.payGroups[index])
            nbo.gEntry['values'] = nbo.payGroups
            nbo.PayGroupVariable.set(NewPayGroupName.get())
            # endregion resetting menus

            self.master.destroy()

        # endregion SQL submittion

    def close_windows(self):
        self.master.destroy()

class cTableWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("View Consolidated Table")
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

def main():
    # region login
    # opens config file
    Config = configparser.ConfigParser()
    Config.read("config.ini")

    # reading base64 login from config.ini
    driver =    Config.get("Login","Driver")
    server =    Config.get("Login","Server")
    database =  Config.get("Login","Database")
    uid =       Config.get("Login","uid")
    pwd =       Config.get("Login","pwd")

    # decoding credentials
    driver =    str(base64.b64decode(driver)).strip('b\'')
    server =    str(base64.b64decode(server)).strip('b\'')
    ind =       server.index("\\\\")
    server =    server[:ind] + server[ind+1:]
    database =  str(base64.b64decode(database)).strip('b\'')
    uid =       str(base64.b64decode(uid)).strip('b\'')
    pwd =       str(base64.b64decode(pwd)).strip('b\'')

    login =     ("Driver=%s;Server=%s;Database=%s;uid=%s;pwd=%s" % (driver, server, database, uid, pwd))

    connection = pypyodbc.connect(login)
    # endregion login

    # region open tkinter window
    root = tk.Tk()
    note = ttk.Notebook(root)
    app = MainApplication(root, note, connection, server)
    root.mainloop()
    connection.close()
    # endregion open tkinter window

if __name__ == '__main__':
    main()