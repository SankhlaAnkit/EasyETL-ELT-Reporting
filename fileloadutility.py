import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter import ttk
from tkinter import font
import re
from io import StringIO
import os
import pyodbc
import pandas as pd
from datetime import datetime
## Main Window ##
FLwindow = tk.Tk()
FLwindow.resizable(False,False)
FLwindow.title('Welcome')
## Frames ##
frame1 = tk.Frame(FLwindow)
frame2 = tk.Frame(FLwindow)
frame3 = tk.Frame(FLwindow)
frame4 = tk.Frame(FLwindow)
frame1.grid()
frame2.grid()
frame3.grid()
frame4.grid()

##Tooltip Class and function for "Hover for info" feature##
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 57
        y = y + cy + self.widget.winfo_rooty() +20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        tiplabel = tk.Label(tw, text=self.text, justify=tk.LEFT,
                      background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        tiplabel.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

## Variable initialization ##
## String ##
dirname=''
fwh=''
fwfile=pd.DataFrame()

## List ## 
tablename=[]
"""
,,path,type1,servername,table,droptable,dbname,headr, \
cols,filedf,conn,cursor,esheet,erow,ecol,fwh,fwcols,All_tables,tablename, \
fwfile,Path,var,var1,var2,var3,frame9,startpos,endpos = ''fw=''
"""


##Functions##
##Called by "Load" Button on Main Window##
def create():
    type1=listbox.get(listbox.curselection())
    if fwbtn3.winfo_exists():
        fw=''
    else:
        fw=var4.get()
    if fw==1 or type1=='xlsx':
        configureSpec()
    else:
        createtable()

##Function called for normal Excel file spec or FW file specs##
def configureSpec():
    global esheet
    global erow
    global ecol
    if fwbtn3.winfo_exists():
        type2=''
    else:
        type2=listbox10.get(listbox10.curselection())    
    type1=listbox.get(listbox.curselection())
    newWindow = tk.Toplevel(FLwindow)
    newWindow.resizable(True, True)
    frame5 = tk.Frame(newWindow)
    frame5.grid()
    label0=tk.Label(frame5,text='Enter Excel file details')
    label0.grid(row=0,column=1)  
    label1=tk.Label(frame5,text='Sheet Name')
    label1.grid(row=1,column=0,padx=10,pady=5)
    esheet = tk.Entry(frame5,width=10)
    esheet.grid(row=1,column=1,padx=10,pady=5)
    
    label2=tk.Label(frame5,text='Skip Rows')
    label2.grid(row=2,column=0,padx=10,pady=5)
    erow = tk.Entry(frame5,width=10)
    erow.grid(row=2,column=1,padx=10,pady=5)
    CreateToolTip(erow, text ='Row number where column configuration begins,Eg. 5')
    label3=tk.Label(frame5,text='Column Range')
    label3.grid(row=3,column=0,padx=10,pady=5)
    ecol = tk.Entry(frame5,width=10)
    ecol.grid(row=3,column=1,padx=10,pady=5)
    CreateToolTip(ecol, text ='Specify the colums with column name and column range,Eg. A:B')
        
    if type2=='xlsx':
        button1=tk.Button(frame5,text='Configure',command=fwload)
        button1.grid(row=4,column=1,padx=10,pady=5)
        CreateToolTip(esheet, text ='Sheet name with column configurations')
    if type1=='xlsx':
        CreateToolTip(esheet, text ='Sheet name with column configurations\n''Specify None for selecting all sheets')
        button1=tk.Button(frame5,text='Load',command=createtable)
        button1.grid(row=4,column=1,padx=10,pady=5)

##Functions for fixedwidth file data load##

### Upload FW files with saved configuration in Fixedwidthinfo.csv file ###
### calls fixwidthload function ###
def fwautoload():
    global fwfile,Path,fwh
    fwh=1
    fwhist=pd.read_csv('fixedwidthinfo.csv',header=0,converters={'Colspecs': eval})
    filename=fwhist['Filename'].tolist()
    fwhist['Showname']=fwhist['Filename'] + ' (' + fwhist['CreateDate'] + ') '
    Showname=fwhist['Showname'].tolist()
    newWindow = tk.Toplevel(FLwindow)
    newWindow.resizable(True, False)
    container = ttk.Frame(newWindow,borderwidth=5,relief="solid")
    frameCanvas = tk.Canvas(container,width=250,height=250)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=frameCanvas.yview)
    scrollable_frame = ttk.Frame(frameCanvas)
    scrollable_frame.bind(
            "<Configure>",
            lambda e: frameCanvas.configure(
            scrollregion=frameCanvas.bbox("all")
          )
        )
    frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
    frameCanvas.configure(yscrollcommand=scrollbar.set)    
    container.grid(row=2,column=1,padx=5,pady=5)
    frameCanvas.pack(side="left")
    scrollbar.pack(side="right", fill="y")
    def chktab():
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        filename=[(a) for a, b in checked if b == 1 ]
        for val in filename:
            if PEntry.get()=='':
                cwd = os.getcwd()
                Path=(cwd + '/' + val).replace('\\', '/')
            else:
                Path=(PEntry.get() + '/' + val).replace('\\', '/')
            
            w=(fwhist.query('Filename==@val')['Colspecs']).tolist()
            cols=(fwhist.query('Filename==@val')['Columns']).tolist()
            fwcols=cols[0].split(',')
            fwfile=pd.read_fwf(Path,header=None,names=fwcols,colspecs=w[0],dtype=str).fillna('')
            fixwidthload(Path,fwh)
            
            
       
    data={}
    for x in filename:
        chkvar = tk.IntVar()
        C1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
        C1.pack(anchor="nw")
        data[x] = chkvar
    PLabel=tk.Label(scrollable_frame,text='FilePath')
    PLabel.pack(anchor="sw")
    PEntry = tk.Entry(scrollable_frame,width=30)
    PEntry.pack(side='right')
    CreateToolTip(PEntry, text ='Specify the path where fixedwidthfile is available\n'
                                 '***Specify Server name and DB name in Main window***')
    
    btn = tk.Button(container, text="Load", command=chktab)
    btn.pack()
    CreateToolTip(btn, text ='Create table and load the selected fixedwidth file\n'
                              'using configuration from "fixedwidthinfo.csv"')
    btn2 = tk.Button(container, text="Close", command=newWindow.destroy)
    btn2.pack()

###Get FW config from Spec file and create FW dataframe,define start/end positions###
def fwload():
    
    global cols
    global startpos
    global endpos
    specfile=e10.get()
    sheetnum=(esheet.get())
    rowskip=int(erow.get())
    colrange=ecol.get()
      
    
    df=pd.read_excel(specfile,sheet_name=sheetnum,dtype=str,header=None,skiprows=range(0,rowskip),
                usecols = colrange,names=['Name', 'Spec','StartPos'])
    df.dropna(inplace = True)
    sub='\('
    df['colB1']=df['Spec'].str.contains(sub)
    df['colF']=df['Spec'].str.extract(r"\((.*?)\)", expand=False).where(df['colB1'] ==True)
    df['colH']=df['Spec'].str.extract(r"V9\((.*?)\)", expand=False).where(df['colB1'] ==True)

    df=df.fillna(0)

    df['colG']=df.colF.astype(int) + df.colH.astype(int)
    df['colG']=df['colG'].replace(0,1)
    
    df['colD']=0

    for ind,row in df.iterrows():
        if ind==0:
            temp=row['colG']
        else:
            df.at[ind,'colD']=temp
            temp+=row['colG']

    df['colH']=df['colD']+df['colG']
    
    cols=list(df['Name'])
    startpos=list(df['colD'])
    endpos=list(df['colH'])
    
    #w=list(zip(df['colD'], df['colH']))
    #w=[(0,3),(3,13),(13,16),(16,25),(25,28),(28,31),(31,None)]
    
    #df1=pd.read_fwf(datfile,header=None,names=col,colspecs=w,dtype=str).fillna('')
    mapcolumns()

###Function to project FW dataframe column start and end position on New window###
def mapcolumns():
        global frame9
        newWindow = tk.Toplevel(FLwindow)
        newWindow.resizable(True, True)
        
        frame8 = tk.Frame(newWindow)
        frame8.config(bd=1, relief=tk.SOLID)
        frame9 = tk.Frame(newWindow)
        frame9.config(bd=1, relief=tk.SOLID)
        frame10=tk.Frame(newWindow)
        
        container = ttk.Frame(frame8)
        frameCanvas = tk.Canvas(container,width=775)
        
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=frameCanvas.yview)

        scrollable_frame = ttk.Frame(frameCanvas)


        scrollable_frame.bind(
            "<Configure>",
            lambda e: frameCanvas.configure(
            scrollregion=frameCanvas.bbox("all")
          )
        )

        frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")

        frameCanvas.configure(yscrollcommand=scrollbar.set)
        
        frame8.pack()
        container.pack(fill="x")
        frameCanvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        button1=tk.Button(container,text='Save/Showdata',command=fixwidthdat)
        button1.pack(side="right")
        CreateToolTip(button1, text ='Show data based on Column configurations on Current window\n'
                                      'Note:-configuration will be saved in "fixedwidthinfo.csv" file\n'
                                       'for future reloads')
        frame9.pack()
        frame10.pack()
        button10=tk.Button(frame10,text='Load',command=fixwidthload)
        button10.pack()
        CreateToolTip(button10, text ='Creates table and Loads file/s into DB\n'
                                     'Note:-for reloading fixedwidth files,Use "fixedwidthhistory"')
        


    
        l1=tk.Label(scrollable_frame,text='Start Position')
        l1.grid(row=0,column=2,ipady=2)
        l2=tk.Label(scrollable_frame,text='End Position')
        l2.grid(row=0,column=3,ipady=2)
        l3=tk.Label(scrollable_frame,text='Start Position')
        l3.grid(row=0,column=6,ipady=2)
        l4=tk.Label(scrollable_frame,text='End Position')
        l4.grid(row=0,column=7,ipady=2)
        
        

        y=1
        global var,var1,var2,var3
        var= []
        var1=[]
        var2=[]
        var3=[]
        mI = (len(cols)-1)//2 

        for x in cols[0:mI+1]:
            l2=tk.Label(scrollable_frame,text=x)
            l2.grid(row=cols.index(x)+1,column=y,padx=5,pady=3)
            e2=tk.Entry(scrollable_frame,width=10)
            e2.grid(row=cols.index(x)+1,column=y+1,padx=5,pady=3)
            var.append(e2)
            e3=tk.Entry(scrollable_frame,width=10)
            e3.grid(row=cols.index(x)+1,column=y+2,padx=5,pady=3)
            var1.append(e3)
        y=5  
        for x in cols[mI+1:]:
            l2=tk.Label(scrollable_frame,text=x)
            l2.grid(row=cols.index(x)-mI,column=y,padx=5,pady=3)
            e2=tk.Entry(scrollable_frame,width=10)
            e2.grid(row=cols.index(x)-mI,column=y+1,padx=5,pady=3)
            var2.append(e2)
            e3=tk.Entry(scrollable_frame,width=10)
            e3.grid(row=cols.index(x)-mI,column=y+2,padx=5,pady=3)
            var3.append(e3)
      
        for i in range(len(var)):
            C=var[i].insert(0,startpos[i])
            C=var1[i].insert(0,endpos[i])
        x=i+1
        for i in range(len(var2)):
            C=var2[i].insert(0,startpos[i+x])
            C=var3[i].insert(0,endpos[i+x])

###Get start/end pos values form new window, create FW DF again and project on new window###
###Save NEW FW config in fixedwidthinfo.csv file for future use###
def fixwidthdat():
    global fwfile,Path
    selct1=[]
    selct3=[]
    for i in range(len(var)):
        C=var[i].get()
        selct1.append(C)
        D=var1[i].get()
        selct3.append(D)
    
    selct2=[]
    selct4=[]
    for i in range(len(var2)):
        C=var2[i].get()
        selct2.append(C)
        D=var3[i].get()
        selct4.append(D)
    
    selct1.extend(selct2)
    selct3.extend(selct4)
    
    selct1 = list(map(int, selct1))
    selct3 = list(map(int, selct3))
    w=list(zip(selct1, selct3))
    
    datfile=listboxx.get(listboxx.curselection())
    Path=dirname + '/' + datfile
    
    Masterdic={'Filename':[datfile],'Columns':[cols],'Colspecs':[w],'Path':[path],'CreateDate':[datetime.now().strftime("%d/%m/%Y")]}
    Masterdf = pd.DataFrame(Masterdic)
    if os.path.isfile('fixedwidthinfo.csv'):
        Masterdf.to_csv('fixedwidthinfo.csv', mode='a', header=not os.path.exists('fixedwidthinfo.csv'),index=False)
    else:
        Masterdf.to_csv('fixedwidthinfo.csv',index=False)
    
    fwfile=pd.read_fwf(Path,header=None,names=cols,colspecs=w,dtype=str).fillna('')
    
    container = ttk.Frame(frame9)
    frameCanvas = tk.Canvas(container,width=1300)
        
    scrollbar = ttk.Scrollbar(container, orient="horizontal", command=frameCanvas.xview)

    scrollable_frame = ttk.Frame(frameCanvas)


    scrollable_frame.bind(
        "<Configure>",
        lambda e: frameCanvas.configure(
        scrollregion=frameCanvas.bbox("all")
          )
        )

    frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")

    frameCanvas.configure(xscrollcommand=scrollbar.set)
        
    
    container.pack(expand = True, fill = "both")    
    frameCanvas.pack(fill="both", expand=True)
    scrollbar.pack( side="bottom",fill="x")
    
    label4=tk.Label(scrollable_frame,text='Data for reference',bg="white")
    label4.pack()
    tree = ttk.Treeview(scrollable_frame, height=10, columns=cols, show='headings')
    tree.pack()
    fontss = font.Font(scrollable_frame)
    
    for col in cols:
        col_width=fontss.measure(col)
        tree.heading(col, text=col)
        tree.column(col, width=col_width, anchor=tk.CENTER)
    limitrows=fwfile.head(15)
    for i in range(len(limitrows)):
        tree.insert('','end', values=list(limitrows.loc[i]))

### CReate FW file table and call insertdata funcion###
def fixwidthload(*args):
    global filedf
    global fwcols
    global servername
    global table
    global dbname
    servername=e2.get()
    dbname=e3.get()
    droptable=dropvar.get()
    
    if servername=='' or dbname=='':
        messagebox.showinfo("Warning","Server/Database Name not specified")
    else:
        conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                      Server=servername,
                      Database=dbname,
                      Trusted_Connection='yes',autocommit=True)
        cursor = conn.cursor()
    
    CurrUsr=os.getlogin()
    filedf=fwfile.assign(InsertDate=str(datetime.now().replace(microsecond=0)),InsertUser=CurrUsr)
    fwcols=(list(filedf.columns))
       
    if e4.get()=='':
        datfile=listboxx.get(listboxx.curselection())
        table=os. path. splitext(datfile)[0]
    else:
        if e4.get()!='':
            table=e4.get()
        else:
            Path=args[0]
            base=os.path.basename(Path)
            table=os.path.splitext(base)[0]
    
    string='  Varchar(255)'
    
    if any(" " in s for s in fwcols):
        A= ','.join(['[' + s + '] ' + string for s in fwcols])
    else:    
        A= ','.join([s+ string for s in fwcols])
    
    X='CREATE TABLE' + ' '+ table + ' (' + A  + ')'

    if droptable==1:
        query1="IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' \
                    AND  TABLE_NAME = '{}') \
                    DROP TABLE {}.dbo.{}".format(table,dbname,table)
        #query1="DROP TABLE {}.dbo.{}".format(dbname,table)
        cursor.execute(query1)
        print("Table dropped and Creating new table")
        cursor.execute(X)
        conn.commit()
        print("New table Created - {}".format(table))
        messagebox.showinfo("Confirmation","Table dropped and New table Created\nData insert in progress..")
        insertdata()
        #button3=tk.Button(frame4,text='Show tables',command=gettables).grid(row=13,column=3)
        #button2=tk.Button(frame4,text='Staging',command=startstgload).grid(row=14,column=3,ipady=5)
    else:
        if cursor.tables(table=table, tableType='TABLE').fetchone():
            insertdata()
            #button3=tk.Button(frame4,text='Show tables',command=gettables).grid(row=13,column=3)
            #button2=tk.Button(frame4,text='Staging',command=startstgload).grid(row=14,column=3,ipady=5)
        else:
            cursor.execute(X)            
            conn.commit()
            print("New table Created - {}".format(table))
            messagebox.showinfo("Confirmation","New table Created\nData insert in progress..")
            insertdata()
            #button3=tk.Button(frame4,text='Show tables',command=gettables).grid(row=13,column=3,ipady=5)
            #button2=tk.Button(frame4,text='Staging',command=startstgload).grid(row=14,column=3,ipady=5)
    


##Function for regular file data load##
##Creates table and calls insertdata function##
def createtable():
    global path
    global type1
    global servername
    global table
    global droptable
    global dbname
    global headr
    global cols
    global filedf
    type1=listbox.get(listbox.curselection())
    servername=e2.get()
    dbname=e3.get()
    tab=e4.get()
    droptable=dropvar.get()
    headr=headvar.get()
    if headr==1:
        headr=None
    
    filename = [listboxx.get(i) for i in listboxx.curselection()]
    if servername=='' or dbname=='':
        messagebox.showwarning("",'Server/Database name not specified')
    else:
        conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                      Server=servername,
                      Database=dbname,
                      Trusted_Connection='yes',autocommit=True)
        cursor = conn.cursor()
    
    CurrUsr=os.getlogin()
    
    for file in filename:
        path=dirname + '/' + file
        ind=filename.index(file)
        
        if tab=='':
            table=os. path. splitext(file)[0]
        else:
            tab1=tab.split(',')
            table=tab1[ind]
            
        if type1=='csv' or type1=='txt':
                try:
                    mainfile=pd.read_csv(path,header=headr)
                    x=0
                except:
                    x=1
                    #print(x)
                if x==1:
                    mainfile=pd.read_csv(path,header=headr,engine='python')
            
        if type1=='xlsx':
            sheetnum=esheet.get()
            if sheetnum=='':
                mainfile=pd.read_excel(path)
            else:
                if erow.get()=='':
                    rowskip=0
                else:
                    rowskip=int(erow.get())
                if ecol.get()=='':
                    mainfile=pd.read_excel(path,sheet_name=sheetnum,dtype=str,header=headr,skiprows=range(0,rowskip))
                else:
                    colrange=ecol.get()
                    mainfile=pd.read_excel(path,sheet_name=sheetnum,dtype=str,header=headr,skiprows=range(0,rowskip),
                    usecols = colrange)
                
        
        
        filedf = mainfile.fillna("null")
        file2=filedf.assign(InsertDate=str(datetime.now().replace(microsecond=0)),InsertUser=CurrUsr)
        filedf = file2.applymap(str)
        cols=(list(file2.columns))
        string='  Varchar(255)'
        string2='Column'
        
        if headr==None:
            if type1=='xlsx':
                A=','.join([string2 + str(s)  + string for s in range(len(cols))])
            else:
                A=','.join([string2+str(s)+string for s in cols])
        else:   
            if any((" " in s or "," in s) for s in cols):
                A= ','.join(['[' + s + '] ' + string for s in cols])
            else:    
                A= ','.join([s+ string for s in cols])
        
    #spacecols=[s for s in cols if " " in s]
        
    
        X='CREATE TABLE' + ' '+ table + '(' + A  + ')'
        
        if droptable==1:
            query1="IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' \
                    AND  TABLE_NAME = '{}') \
                    DROP TABLE {}.dbo.{}".format(table,dbname,table)
            #query1="DROP TABLE {}.dbo.{}".format(dbname,table)
            cursor.execute(query1)
            print("Table dropped and Creating new table")
            cursor.execute(X)
            conn.commit()
            print("New table Created - {}".format(table))
            messagebox.showinfo("Confirmation","Table dropped and New table Created\nData insert in progress..")
            insertdata()
            #button3=tk.Button(frame4,text='Show tables',command=gettables).grid(row=13,column=4)
            #button2=tk.Button(frame4,text='Staging',command=startstgload).grid(row=14,column=4,ipady=5)
        else:
            if cursor.tables(table=table, tableType='TABLE').fetchone():
                insertdata()
                #button3=tk.Button(frame4,text='Show tables',command=gettables).grid(row=13,column=4)
                #button2=tk.Button(frame4,text='Staging',command=startstgload).grid(row=14,column=4,ipady=5)
            else:
                cursor.execute(X)            
                conn.commit()
                print("New table Created - {}".format(table))
                messagebox.showinfo("Confirmation","New table Created\nData insert in progress..")
                insertdata()
                #button3=tk.Button(frame4,text='Show tables',command=gettables).grid(row=13,column=4,ipady=5)
                #button2=tk.Button(frame4,text='Staging',command=startstgload).grid(row=14,column=4,ipady=5)

    path=''
    headr=''
    type1=''
    servername=''
    table=''
    droptable=''
    dbname=''

## Common Functions ##
### Function to inserdata - Regular & fixedwidth file ###
def insertdata(*args):
    global conn
    global cursor
    conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                      Server=servername,
                      Database=dbname,
                      Trusted_Connection='yes',automcommit=True)
    cursor = conn.cursor()
    
    query='select * from ' + table
    string2='Column'
    for i, col in enumerate(filedf.columns):
            filedf.iloc[:, i] = filedf.iloc[:, i].str.replace("'", "''")
            filedf.iloc[:, i] = filedf.iloc[:, i].str.replace('"', "''")
    if fwbtn3.winfo_exists():
        fw=''
    else:
        fw=var4.get()
    if fw==1:
        #file = fixwidthDf.applymap(str)
        B= ','.join(['[' + s + '] ' for s in fwcols])
    
    if fwh==1:
        #file = fixwidthDf.applymap(str)
        B= ','.join(['[' + s + '] ' for s in fwcols])
    else:
        if headr==None:
            if type1=='xlsx':
                B=','.join([string2 + str(s) for s in range(len(cols))])
            else:
                B=','.join([string2+str(s) for s in cols])
        else:   
            if any((" " in s or "," in s) for s in cols):
                B= ','.join(['[' + s + '] ' for s in cols])
            else:    
                B= ','.join([s for s in cols])
    
    records = [str(tuple(x)) for x in filedf.values]
    
    insrt = 'Insert into ' + table + ' ' + '(' + B  + ')' + ' values '

    def chunker(seq, size):
        return (seq[pos:pos + size] for pos in range(0, len(seq), size))
    if cursor.execute(query).rowcount==0:
        for batch in chunker(records, 1000):
            #rows=str(batch).strip(‘[]’)
            rows = ','.join(batch)            
            insertrows = insrt + rows
            insertrows=(re.sub('"',"'", insertrows))
            cursor.execute(insertrows)
            conn.commit()
        print("Data inserted in table - {}".format(table))
        messagebox.showinfo("Confirmation","Data inserted in Table")
    else:
        result=messagebox.askquestion("Confirmation",'Data exists in table\n'
                                          'Do you want to append?')
        if result=='yes':   
            for batch in chunker(records, 1000):
                 rows = ','.join(batch)            
                 insertrows = insrt + rows
                 insertrows=(re.sub('"',"'", insertrows))
                 cursor.execute(insertrows)
                 conn.commit()
            print("Data Append in table - {} Successful".format(table))
            messagebox.showinfo("Confirmation","Data Appened in Table,Successful")

### Function to get all tables from DB ###
def gettables():
    global servername,dbname,conn,cursor
    servername=e2.get()
    dbname=e3.get()
    if servername=='' or dbname=='':
        messagebox.showwarning("Warning","Server/Database Name not specified")
    else:
        conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                      Server=servername,
                      Database=dbname,
                      Trusted_Connection='yes',automcommit=True)
        cursor = conn.cursor()
    
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=?",dbname)
    All_tables=sorted([item[0] for item in cursor.fetchall()],key=str.lower)
    newWindow = tk.Toplevel(FLwindow)
    newWindow.resizable(True, False)
    container = ttk.Frame(newWindow,borderwidth=5,relief="solid")
    frameCanvas = tk.Canvas(container,width=180,height=250)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=frameCanvas.yview)
    scrollable_frame = ttk.Frame(frameCanvas)
    scrollable_frame.bind(
            "<Configure>",
            lambda e: frameCanvas.configure(
            scrollregion=frameCanvas.bbox("all")
          )
        )
    frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
    frameCanvas.configure(yscrollcommand=scrollbar.set)    
    container.grid(row=2,column=1,padx=5,pady=5)
    frameCanvas.pack(side="left")
    scrollbar.pack(side="right", fill="y")
    
    filterlist=[]
    def find():
       cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=?",dbname)
       All_tables=sorted([item[0] for item in cursor.fetchall()],key=str.lower)
       s = edit.get()
       if s=='':
           matches = [x for x in All_tables]
       else:
           matches = [x for x in All_tables if s in x]
       for i in check_box_list:
            i.forget()
       for i in filterlist:
            i.forget()
       for x in matches:
           chkvar = tk.IntVar()
           D1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
           D1.pack(anchor="nw")
           data[x] = chkvar
           filterlist.append(D1)
    
    edit = tk.Entry(scrollable_frame) 
  
#positioning of text box
    edit.pack(fill="x", expand=1) 
  
#setting focus
    edit.focus_set() 
  
#adding of search button
    butt = tk.Button(scrollable_frame, text='Find',command=find)  
    butt.pack()
  
    
    def chktab():
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        tablename=[(a) for a, b in checked if b == 1 ]
        for val in tablename:
            getdata(val)
            
    def droptab():
        global All_tables
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        tablename=[(a) for a, b in checked if b == 1 ]
        for val in tablename:
            query1="DROP TABLE {}.dbo.{}".format(dbname,val)
            cursor.execute(query1)
            print(val + " Table dropped")
            data.pop(val)
        messagebox.showinfo("Confirmation","All Selected Tables dropped")
        cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=?",dbname)
        All_tables=sorted([item[0] for item in cursor.fetchall()],key=str.lower)
        
        for i in check_box_list:
            i.forget()
        for i in filterlist:
            i.forget()
        if edit.get()=='':
            
            for x in All_tables:
                chkvar = tk.IntVar()
                C1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
                C1.pack(anchor="nw")
                check_box_list.append(C1)
                data[x] = chkvar
        else:
            matches = [x for x in All_tables if edit.get() in x]
            for x in matches:
                chkvar = tk.IntVar()
                D1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
                D1.pack(anchor="nw")
                data[x] = chkvar
                filterlist.append(D1)

    data={}
    check_box_list=[]
    for x in All_tables:
        chkvar = tk.IntVar()
        C1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
        C1.pack(anchor="nw")
        check_box_list.append(C1)
        data[x] = chkvar
    
    btn = tk.Button(container, text="Data", command=chktab)
    btn.pack()
    CreateToolTip(btn, text ='Shows top 10 rows of selected tables')
    btn1 = tk.Button(container, text="Drop", command=droptab)
    btn1.pack()
    CreateToolTip(btn1, text ='drops selected tables')
    btn2 = tk.Button(container, text="Close", command=newWindow.destroy)
    btn2.pack()


### Function to get data for selected tables ###  
def getdata(*args):
    for val in (args or tablename):
        top10lnd='Select top 10 * from {}'.format(val)
        lnddf = pd.read_sql_query(top10lnd, conn)
        lndcols=list(lnddf.columns)
        newWindow = tk.Toplevel(FLwindow)
        newWindow.resizable(True, False)
        #container = ttk.Frame(frame9)
        frameCanvas = tk.Canvas(newWindow,width=1000)
        scrollbar = ttk.Scrollbar(newWindow, orient="horizontal", command=frameCanvas.xview)
        scrollable_frame = ttk.Frame(frameCanvas)
        scrollable_frame.bind(
        "<Configure>",
        lambda e: frameCanvas.configure(
        scrollregion=frameCanvas.bbox("all")
          )
        )
        frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
        frameCanvas.configure(xscrollcommand=scrollbar.set)
        #container.pack(expand = True, fill = "both")    
        frameCanvas.pack(fill="both", expand=True)
        scrollbar.pack( side="bottom",fill="x")
        
        #newCanv = tk.Canvas(newWindow, bg="white")
        #newCanv.pack()
        label5=tk.Label(scrollable_frame,text='{} table data for reference'.format(val))
        label5.grid(row=0,column=0)
        tree = ttk.Treeview(scrollable_frame, height=10, columns=lndcols, show='headings')
        tree.grid(row=1, column=0, sticky='news',padx=10)
        fontss = font.Font(scrollable_frame)
        for col in lndcols:            
            col_width=fontss.measure(col)
            tree.heading(col, text=col)
            tree.column(col, width=col_width, anchor=tk.CENTER)
        for i in range(len(lnddf)):
            tree.insert('','end', values=list(lnddf.loc[i]))
        """
        sb = tk.Scrollbar(scrollable_frame, orient=tk.VERTICAL, command=tree.yview)
        sb.grid(row=1, column=1, sticky='ns')
        tree.config(yscrollcommand=sb.set)
        """
        button6=tk.Button(scrollable_frame,text='Close data',command= newWindow.destroy)
        button6.grid(row=2,column=0)
    

### Function to clear all field values on Main window ###
def Clearall():
    listboxx.delete(0,tk.END)
    listbox.selection_clear(0, 'end')
    listbox10.selection_clear(0, 'end')
    e2.delete(0,tk.END)
    e3.delete(0,tk.END)
    e4.delete(0,tk.END)
    e10.delete(0,tk.END)
    chbox1.deselect()
    chbox5.deselect()
    chbox4.deselect()


"""
##Error display##
def report_callback_exception(self, exc, val, tb):
    if 'must be active, anchor, end, @x,y, or a number' in str(val):
        messagebox.showerror("Error", message='No Value Selected from List')
    if "'cursor' referenced before assignment" in str(val):
        messagebox.showerror("Error", message='Server/Database name not specified')
    else:
        messagebox.showerror("Error", message=str(val))

tk.Tk.report_callback_exception = report_callback_exception
"""
### Function to call staging load file ###
def startstgload():
    Count=1
    dbname=e3.get()
    servername=e2.get()
    #%store dbname
    #%store servername
    #exec(open("AStoredProc&TableLoad.ipynb").read())
    #os.system("exec AStoredProc&TableLoad.ipynb")
    #os.system('python AStoredProc&TableLoad.ipynb')
    result=messagebox.askquestion("Confirmation", 'Do you want to Close\n'
                                          'File to table load window?')
    if result=='yes':   
        FLwindow.destroy()
        
    #%run ./AStoredProc&TableLoad.ipynb

## Main Window Design Components ##

MainHeading= tk.Label(frame1,text="Please Provide File Details",justify=tk.CENTER)
MainHeading.grid(row=1,column=2, ipady=5)
listbox10 =tk.Listbox(frame3,height=3,  
                  width = 20,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",12), 
                  fg = "black",
                  exportselection=False,selectmode = "multiple")
e10 = tk.Entry(frame3,width=60)
chbox5=tk.Checkbutton(frame2,text='Fixed width')

def listfiles():
    global dirname
    listboxx.delete(0,tk.END)
    dirname = fd.askdirectory()
    filelist=[f for f in os.listdir(dirname) if (f.endswith('.txt') or f.endswith('.csv') or f.endswith('.xml') or f.endswith('.xlsx'))]
    mylist=sorted(filelist,key=str.lower)
    for file in mylist:
        listboxx.insert(tk.END, file)

label1=tk.Label(frame2,text='Select files for Profiling')
label1.grid(row=2)
listboxx =tk.Listbox(frame2,height=5,  
                  width = 40,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",
                  exportselection=False,selectmode = "multiple")
listboxx.grid(row=2,column=2)
scrollbar = tk.Scrollbar(frame2)
scrollbar.grid(row = 2,rowspan = 1, column = 3,ipady=15)
listboxx.config(yscrollcommand = scrollbar.set)
scrollbar.config(command = listboxx.yview)
button1=tk.Button(frame2,text='...',command=listfiles)
button1.grid(row=2,column=4,padx=10)

CreateToolTip(button1, text = 'Select directory')

headvar = tk.IntVar()

chbox1=tk.Checkbutton(frame2,text='No header',variable=headvar)
chbox1.grid(row=3,column=2)
CreateToolTip(chbox1, text = 'Select if file header is not defined')

FTlabel=tk.Label(frame2,text='File Type')
FTlabel.grid(row=5)
listbox =tk.Listbox(frame2,height=3,  
                  width = 20,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",12), 
                  fg = "black",
                  exportselection=False,selectmode = "multiple")
listbox.grid(row=5,column=2,pady=5)
scrollbar = tk.Scrollbar(frame2)
scrollbar.grid(row = 5,rowspan = 1, column = 3,ipady=15,pady=5)
listbox.config(yscrollcommand = scrollbar.set)
scrollbar.config(command = listbox.yview)

### Main window display functions ###


def callback():
    e10.delete(0, "end")
    currdir=os.getcwd()
    name= fd.askopenfilename(initialdir=currdir)
    e10.insert(0,name)

def askfwspec():
    global e10,listbox10,chbox5,var4
    #fw,
    fwbtn3.destroy()
    label10=tk.Label(frame3,text='Configuration File')
    label10.grid(row=6)
    CreateToolTip(label10, text = 'file that defines the columns\n'
                               'and width of fixed width file')
    e10 = tk.Entry(frame3,width=60)
    e10.grid(row=6, column=2,padx=5)
    button10=tk.Button(frame3,text='...',command=callback)
    button10.grid(row=6,column=3)
    CreateToolTip(button10, text = 'select specification file')

    SPFTlabel=tk.Label(frame3,text='Spec File Type')
    SPFTlabel.grid(row=8)
    listbox10 =tk.Listbox(frame3,height=3,  
                  width = 20,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",12), 
                  fg = "black",
                  exportselection=False,selectmode = "multiple")
    listbox10.grid(row=8,column=2,pady=5)
    scrollbar10 = tk.Scrollbar(frame3)
    scrollbar10.grid(row = 8,rowspan = 1, column = 3,ipady=15,pady=5)
    listbox10.config(yscrollcommand = scrollbar10.set)
    scrollbar10.config(command = listbox10.yview)
    
    for i in typelist:
        listbox10.insert(tk.END,i)
    var4 = tk.IntVar()
    chbox5=tk.Checkbutton(frame2,text='Fixed width',variable=var4)
    chbox5.grid(row=4,column=2)
    CreateToolTip(chbox5, text = 'Select if source is fixed width')
    chbox5.select()

fwbtn3=tk.Button(frame2,text='Fixed Width',command=askfwspec)
fwbtn3.grid(row=4,column=2)
CreateToolTip(fwbtn3, text = 'Select if source is fixed width')

typelist=['csv','txt','xml','xlsx']
for i in typelist:
    listbox.insert(tk.END,i)
    
Servlabel=tk.Label(frame4,text='Server Name')
Servlabel.grid(row=9,padx=50)
e2 = tk.Entry(frame4,width=60)
e2.grid(row=9, column=2,pady=2)
CreateToolTip(e2, text = 'Name of Database Server\n'
                        'by default,Windows creds  will be used')

DBlabel=tk.Label(frame4,text='Database Name')
DBlabel.grid(row=10)
e3 = tk.Entry(frame4,width=60)
e3.grid(row=10, column=2,pady=2)
CreateToolTip(e3, text = 'Name of Database Schema')

Tablabel=tk.Label(frame4,text='Table Name')
Tablabel.grid(row=11)
e4 = tk.Entry(frame4,width=60)
e4.grid(row=11, column=2,pady=2)
CreateToolTip(e4, text = 'Specify table name for each file selected\n'
                          'Comma separated values.Eg. Tab1,Tab2..\n'
                         'Leave blank to set filename as table name')

dropvar = tk.IntVar()
chbox4=tk.Checkbutton(frame4,text='drop table',variable=dropvar)
chbox4.grid(row=13,column=2)
CreateToolTip(chbox4, text = 'select if table exists and should be re-created\n'
                         'Note:if mutiple file/table selected, all will be dropped\n'
                         '      You can use "Show tables" to drop tables')

## Common Buttons ##
button1=tk.Button(frame4,text='Load',command=create)
button1.grid(row=13,column=0,pady=2)
CreateToolTip(button1, text ='Creates table and Loads file/s into DB\n'
                             'Note:-for reloading fixedwidth files,Use "fixedwidthhistory"')
button6=tk.Button(frame4,text='ClearAll',command=Clearall)
button6.grid(row=14,column=0,pady=2)
CreateToolTip(button6, text ='Clear All entries and selection')
Destroybutton=tk.Button(frame4,text='Close',command=FLwindow.destroy)
Destroybutton.grid(row=15,column=0,pady=5)
button3=tk.Button(frame4,text='Show tables',command=gettables)
button3.grid(row=13,column=3,pady=2)
CreateToolTip(button3, text = 'List all tables in DB Schema\n'
                          'Drop tables, View table data\n'
                         '***Specify Server name and DB name***')
button4=tk.Button(frame4,text='Staging',command=startstgload)
button4.grid(row=14,column=3,pady=2)
CreateToolTip(button4, text = 'Create SPs for table to table load')
button5=tk.Button(frame4,text='fixedwidthhistory',command=fwautoload)
button5.grid(row=15,column=3,pady=5,padx=10)
CreateToolTip(button5, text ='Reload Fixed width files\n'
                             'Note:-Make sure "fixedwidthinfo.csv" is available in current environment\n'
                             '***Specify Server name and DB name***')

FLwindow.mainloop()