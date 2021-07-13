import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import filedialog
from tkinter import messagebox
from collections import Counter
import os
import pyodbc
import pandas as pd
import pandas.io.formats.excel
pandas.io.formats.excel.ExcelFormatter.header_style = None
import re
from datetime import datetime
from pandas import ExcelWriter
import xlsxwriter
import string
window = tk.Tk()
#window.geometry("1200x800")
window.state('zoomed')
window.configure(bg="white")
window.title('Welcome')
myCanvas = tk.Canvas(window, bg="white")
myCanvas2 = tk.Canvas(window, bg="white",highlightbackground="black",highlightthickness=1)
frame13 = tk.Frame(myCanvas2,highlightbackground="black",highlightthickness=1)
frame1 = tk.Frame(myCanvas,bg="white")
frame2 = tk.Frame(myCanvas,bg="white")
frame9 = tk.Frame(frame2,bg="white")
frame3 = tk.Frame(myCanvas,bg="white")
frame10 = tk.Frame(frame3,bg="white")
frame11 = tk.Frame(frame3,bg="white")
frame12 = tk.Frame(frame3,bg="white")
frame4 = tk.Frame(myCanvas,bg="white")
frame5 = tk.Frame(myCanvas,bg="white")
frame6 = tk.Frame(frame3,bg="white")
frame7 = tk.Frame(myCanvas,bg="white")
frame8 = tk.Frame(myCanvas,bg="white",highlightbackground="black",highlightthickness=1)


class DragDropListbox(tk.Listbox):
    def __init__(self, master, **kw):
        kw['selectmode'] = tk.MULTIPLE
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.getState, add='+')
        self.bind('<Button-1>', self.setCurrent, add='+')
        self.bind('<B1-Motion>', self.UpDownSelection)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonRelease-1>", self.on_drop)
        self.bind('<Double-Button-1>', self.removeSelection)
        self.configure(cursor="hand1")
        self.curIndex = None

    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)
    def getState(self, event):
        i = self.nearest(event.y)
        self.curState = self.selection_includes(i)
    def removeSelection(self,event):
        self.selection_clear(0, tk.END)

    def UpDownSelection(self, event):
        i = self.nearest(event.y)
        if self.curState == 1:
           self.selection_set(self.curIndex)
        else:
           self.selection_clear(self.curIndex)
        if i < self.curIndex:
           selected = self.selection_includes(i)
           if selected:
               self.selection_set(i+1)
           self.curIndex = i
        if i > self.curIndex:
        #Moves down
           selected = self.selection_includes(i)
           if selected:
               self.selection_set(i-1)
           self.curIndex = i
        
    
    def on_leave(self, event):
        global value
        value = list()
        selection=event.widget.curselection()
        #self.selection_clear(self.curIndex)
        for i in selection:
            entrada = event.widget.get(i)
            value.append(entrada)
        
    
    def on_drop(self, event):
        i='--'
        # find the widget under the cursor
        x,y = event.widget.winfo_pointerxy()
        target = event.widget.winfo_containing(x,y)
        try:
            if isinstance(target,shiftSelectListbox):
                for val in value:
                    if i in val:
                        idx=value.index(val)
                        value.pop(idx)
                for val in value:
                    target.insert(tk.END,val)
        except:
            pass

class shiftSelectListbox(tk.Listbox):
    def __init__(self, master, **kw):
        kw['selectmode'] = tk.SINGLE
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.setCurrent)
        self.bind('<B1-Motion>', self.shiftSelection)
        self.bind('<Double-Button-1>', self.removeSelection)
        self.configure(cursor="hand1")
        self.curIndex = None

    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)

    def shiftSelection(self, event):
        i = self.nearest(event.y)
        if i < self.curIndex:
            val = self.get(i)
            self.delete(i)
            self.insert(i+1, val)
            self.curIndex = i
        elif i > self.curIndex:
            val = self.get(i)
            self.delete(i)
            self.insert(i-1, val)
            self.curIndex = i

    def removeSelection(self,event):
        i = self.nearest(event.y)
        self.delete(i)

def SerDetail():
    global Serve2,ServDetail,conn,cursor
    ServDetail = ttk.Frame(frame13,borderwidth=5,relief="solid")
    frameCanvas = tk.Canvas(ServDetail,width=150)    
    scrollable_frame = ttk.Frame(frameCanvas)
    scrollable_frame.bind(
            "<Configure>",
            lambda e: frameCanvas.configure(
            scrollregion=frameCanvas.bbox("all")
          )
        )
    frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
    frameCanvas.configure(yscrollcommand=scrollbar.set)    
    ServDetail.grid(row=1,column=1,padx=5,pady=5)
    frameCanvas.pack(side="left")

    l2=tk.Label(frameCanvas,text='Server Name')
    l2.grid(row=2,padx=5,pady=5)
    Serve2 = tk.Entry(frameCanvas,width=20)
    Serve2.grid(row=2, column=2,padx=5,pady=5)
    Serve2.insert(tk.END, 'localhost')
    l3=tk.Label(frameCanvas,text='Database Name')
    l3.grid(row=3,padx=5,pady=5)
    servbut= tk.Button(frameCanvas,text="Submit",command=SetServerdet)
    servbut.grid(row=4, column=2,padx=5,pady=5)
    closebut= tk.Button(frameCanvas,text="Close",command=ServDetail.destroy)
    closebut.grid(row=4,padx=5,pady=5)    
    menubutton = tk.Menubutton(frameCanvas,indicatoron=True, borderwidth=1, relief="raised",bg='white')
    menu = tk.Menu(menubutton, tearoff=False,bg='white')
    menubutton.configure(menu=menu)
    menubutton.grid(row=3, column=2,padx=5,pady=5)

    for choice in Databases:
            choices[choice] = tk.IntVar(value=0)
            menu.add_checkbutton(label=choice, variable=choices[choice])

def SetServerdet():
    servername=Serve2.get()
    ServDetail.destroy()
    reports()
    listbox.delete(0,tk.END)
    for name, var in choices.items():
        if var.get()==1:
            dbname.append(name)
            conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                      Server=servername,
                      Database=name,
                      Trusted_Connection='yes',autocommit=True)
            #cursor = conn.cursor()
            Query='select schema_name(tab.schema_id) as schema_name, \n' \
                   'tab.name as table_name, \n' \
                   'col.name as column_name,t.name as data_type,col.max_length,col.precision \n' \
                   'from {}.sys.tables as tab \n' \
                   'inner join {}.sys.columns as col \n' \
                   'on tab.object_id = col.object_id \n' \
                   'left join {}.sys.types as t \n' \
                   'on col.user_type_id = t.user_type_id \n' \
                   'order by schema_name,table_name,column_id;'.format(name,name,name)
            tabcoldf=pd.read_sql_query(Query,conn)
            lnd_tables=tabcoldf['table_name'].unique()
            for i in lnd_tables:
                listbox.insert(tk.END,i)
            for i in lnd_tables:
                df=tabcoldf.loc[tabcoldf['table_name']==i, ['column_name']]
                lnd_columns=df['column_name'].tolist()             
                DB_Table[i]=[name,lnd_columns]

            for value in DB_Table.values():
                selcolms.extend(value[1])
            
    availcols.extend([v + '(' + str(selcolms[:i].count(v) + 1) + ')' if selcolms.count(v) > 1 else v for i, v in enumerate(selcolms)])


def getlndcolumns():
    
    global joincondition,jointype,Wherecondition1,tablename,listbox2,tabindex,colvar

    label2=tk.Label(frame10,text='Source Columns   ',bg="white")
    label2.grid(row=2,ipadx=10)
    listbox2 =DragDropListbox(frame10,height=30,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",exportselection=0,selectmode = "multiple")
    listbox2.grid(row=2,column=2)
    scrollbar2 = tk.Scrollbar(frame10)
    scrollbar2.grid(row = 2,rowspan = 1, column = 3,ipady=15)
    listbox2.config(yscrollcommand = scrollbar2.set)
    scrollbar2.config(command = listbox2.yview)
    button5=tk.Button(frame10,text='Show data',command=getdata)
    button5.grid(row=3,column=2,pady=10)
    seltables = [listbox.get(i) for i in listbox.curselection()]
    tabindex=[]
    tablename=[]
    for val in seltables:
        for key, value in DB_Table.items():        
            if val == key:
                #tablename.append(value[0] + '.dbo.' + val)
                tabname="----" + val + "----"
                listbox2.insert(tk.END,tabname)
                index = listbox2.get(0, "end").index(tabname)
                tabindex.append(index)
                listbox2.itemconfig(index, bg='yellow')
                for col in value[1]:
                    listbox2.insert(tk.END,col)
    
    #for val in tablename:
    #    top10lnd='Select top 10 * from {}'.format(val)
    #    lnddf = pd.read_sql_query(top10lnd, conn)
    #    cols=sorted(list(lnddf.columns),key=str.lower)
    #    tabname="----" + val + "----"
    #    listbox2.insert(tk.END,tabname)
    #    index = listbox2.get(0, "end").index(tabname)
    #    tabindex.append(index)
    #    listbox2.itemconfig(index, bg='yellow')
    #    for col in cols:
    #        listbox2.insert(tk.END,col)
    for widget in frame11.winfo_children():
        widget.destroy()
    colvar=[]
    for i in range(len(tablename)):
        l2=tk.Label(frame11,text=tablename[i],bg="white")
        l2.grid(row=i,column=1,padx=5,pady=3)
        Collistbox =shiftSelectListbox(frame11,height=4,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",exportselection=0)
        Collistbox.grid(row=i,column=2)
        colvar.append(Collistbox)
        scrollbar4 = tk.Scrollbar(frame11)
        scrollbar4.grid(row = i,rowspan = 1, column = 3,ipady=15)
        Collistbox.config(yscrollcommand = scrollbar4.set)
        scrollbar4.config(command = Collistbox.yview)

        l4=tk.Label(frame12,text=' Where Clause',bg="white")
        l4.grid(row=5,column=5,padx=5,pady=3)
        Wherecondition1=tk.Text(frame12,width=25,height=3)
        Wherecondition1.grid(row=5,column=6,padx=5,pady=3,ipady=10)
        sb9 = tk.Scrollbar(frame12, orient=tk.VERTICAL, command=Wherecondition1.yview)
        sb9.grid(row=5,column=7,ipady=3)
        Wherecondition1.config(yscrollcommand=sb9.set)
        
        if len(listbox.curselection())>1:
            l4=tk.Label(frame12,text=' Join Type',bg="white")
            l4.grid(row=3,column=5,padx=5,pady=3)
            jointype=tk.Text(frame12,width=25,height=1.5)
            jointype.grid(row=3,column=6,padx=5,pady=3)
            
            l3=tk.Label(frame12,text=' Join Condition',bg="white")
            l3.grid(row=4,column=5,padx=5,pady=3)
            joincondition=tk.Text(frame12,width=25,height=3)
            joincondition.grid(row=4,column=6,padx=5,pady=3,ipady=10)
            sb9 = tk.Scrollbar(frame12, orient=tk.VERTICAL, command=joincondition.yview)
            sb9.grid(row=4,column=7,ipady=3)
            joincondition.config(yscrollcommand=sb9.set)
        
        frbutton=tk.Frame(frame12,bg='white')
        frbutton.grid(row=6,column=6,ipady=3)
        button6=tk.Button(frbutton,text='Verify',command= customreportsql)
        button6.pack(side='left')
        button7=tk.Button(frbutton,text='Customize',command= Customizereport)
        button7.pack(side='left')

def getdata(*args):
    for val in (args or tablename):
        top10lnd='Select top 10 * from {}'.format(val)
        lnddf = pd.read_sql_query(top10lnd, conn)
        lndcols=list(lnddf.columns)
        newWindow = tk.Toplevel(window)
        newWindow.resizable(True, False)
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
        frameCanvas.pack(fill="both", expand=True)
        scrollbar.pack( side="bottom",fill="x")

        label3=tk.Label(scrollable_frame,text='{} table data for reference'.format(val))
        label3.grid(row=0,column=0)
        tree = ttk.Treeview(scrollable_frame, height=10, columns=lndcols, show='headings')
        tree.grid(row=1, column=0, sticky='news',padx=10)
        fontss = font.Font(scrollable_frame)
        for col in lndcols:            
            col_width=fontss.measure(col)
            tree.heading(col, text=col)
            tree.column(col, width=col_width, anchor=tk.CENTER)
        for i in range(len(lnddf)):
            tree.insert('','end', values=list(lnddf.loc[i]))
        button6=tk.Button(scrollable_frame,text='Close data',command= newWindow.destroy)
        button6.grid(row=2,column=0)

def reports(): 
    listbox3.delete(0,tk.END)
    for i in existingreport:
        listbox3.insert(tk.END,i)
    listbox4.delete(0,tk.END)
    for i in existingderivedcol:
        listbox4.insert(tk.END,i) 

def customreportsql(*args):
    global EData
    EData=[]
    Wh=Wherecondition1.get("1.0","end-1c")
    JoinT=jointype.get("1.0","end-1c")
    Cnd=joincondition.get("1.0","end-1c")
    EReport={}
    colRpt={}
    rowRpt={}
    tabcoldict={}
    finalfrmcmd=[]
    ### where clause ###
    if len(Wh)!=0:
        Wh=Wh.replace(',','\nand\n')
        if ';' in Wh:
            Wh=Wh.split(';')
    
    ### Join Clause ###
    Jtype=JoinT.split(',')
    if len(Cnd)==0:
        JCondition=''
    else:
        CN=Cnd.replace(',','\nand ')
        if ';' in CN:
            JCondition=CN.split(';')
        else:
            JCondition=CN
            
    fromcnd=[]
    for i in range(len(tablename)-1):
        if i==0:
            tab1=tablename[i] + ' tab' + str(i+1) + '\n' 
            Join1=Jtype[i] + ' join '
            tab2=tablename[i+1] + ' tab' + str(i+2) + '\n'
            if len(tablename)==2:
                Jcon='ON ' + JCondition + '\n'
            else:
                Jcon='ON ' + JCondition[i] + '\n'
            fromcnd.append(tab1 + Join1 + tab2 + Jcon)
            
        else:
            Join1=Jtype[i] + ' join '
            tab1=tablename[i+1] + ' tab' + str(i+2) + '\n'
            Jcon='ON ' + JCondition[i] + '\n'
            fromcnd.append(Join1 + tab1 + Jcon)
        
    finalfrmcmd=' '.join(i for i in fromcnd)
    
    ### SQL creation for Report ###
    ## per table Col report without join/where conditions ##
    if len(Wh)==0 and len(JoinT)==0:
        print('Ankit Sankhla')
    ## per table Col report without join , with/without where conditions ##
    if len(Wh)!=0 and len(JoinT)==0:
        for i in range(len(colvar)):
            tabcoldict[tablename[i]]=[colvar[i].get(0,tk.END)]
            if len(Wh)!=0:
                tabcoldict[tablename[i]].append(Wh[i])
        for key,value in tabcoldict.items():
            if len(value[0])==0:
                selcol='*'
            else:
                selcol=','.join(value[0])
            if len(value[1])==0:
                whereclause='' 
            else:
                whereclause='Where ' + ''.join(value[1])

            sql ='Select\n' + selcol + '\nfrom\n' + key  + '\n' + whereclause
            lnddf = pd.read_sql_query(sql,conn)
            rowRpt[key] = lnddf.shape[0]  # Gives number of rows
            colRpt[key] = lnddf.shape[1]  # Gives number of col
            EReport[key]=lnddf
            EData.append(sql)
        print(EData)
        print("col specific file")
    
    ## All table report with join and with/without where conditions ## 
    if len(JoinT)!=0:
        vallist=[]
        wherelist=[]       
        for i in range(len(colvar)):
            tabcoldict[tablename[i]]=[colvar[i].get(0,tk.END)]
            if len(Wh)!=0:
                tabcoldict[tablename[i]].append(Wh[i])
        for key,value in tabcoldict.items():
            vallist.append(value[0])
            if len(Wh)!=0:
                wherelist.append(value[1])
            if len(vallist)==0:
                selcol='*'
            else:
                collist= [','.join(tups) for tups in vallist]
                collist=[x for x in collist if x]
                selcol=','.join(collist)
        if len(wherelist)==0:
            whereclause=''
        else:
            whereclause=[''.join(tups) for tups in wherelist]
            whereclause='Where ' + '\nand '.join(whereclause)
        EData ='Select\n' + 'top 10 ' + selcol + '\nfrom\n' + finalfrmcmd  + '\n' + whereclause
        print("col specific file")
    showsql(EData)  

def showsql(*args):
    
    def savereport(*args):
        if len(RportNam.get())==0:
            messagebox.showinfo("Warning","Please specify Report Name")
        else:
            newrptname=RportNam.get()
            if newrptname in globaldict.keys():
                messagebox.showinfo("Warning","Report Name already exists \n Please specify a different Name")
            else:
                globaldict[newrptname]=EData
                existingreport.append(newrptname)
                Masterdict={'Name':[newrptname],'SQL':[EData],'Type':['Rpt']}
                Masterdf=pd.DataFrame(Masterdict)
                if os.path.isfile(customfilepath):
                    Masterdf.to_csv(customfilepath, mode='a', header=not os.path.exists(customfilepath),index=False)
                else:
                    Masterdf.to_csv(customfilepath,index=False)


    newWindow = tk.Toplevel(window)
    newCanv = tk.Canvas(newWindow, bg="white")
    newCanv.pack(expand=True)
    textbox=tk.Text(newCanv,bg="white",height = 20, width = 100)
    textbox.grid(row=1,column=0,padx=50,pady=50)
    
    sb = tk.Scrollbar(newCanv, orient=tk.VERTICAL, command=textbox.yview)
    sb.grid(row=1,column=1,ipady=50)
    textbox.config(yscrollcommand=sb.set)
    RportNam=tk.Entry(newCanv,width=30)
    RportNam.grid(row=2,column=0,ipady=5)
    button1=tk.Button(newCanv,text='Save Report',command= savereport)
    button1.grid(row=3,column=0,ipady=5)
    button2=tk.Button(newCanv,text='Close',command= newWindow.destroy)
    button2.grid(row=4,column=0,ipady=5)
    button3=tk.Button(newCanv,text='Show data',command=newWindow.destroy)
    button3.grid(row=5,column=0,ipady=5)
    EData=args[0]
    textbox.insert(tk.END,EData) 

def Customizereport(*args):   
    def derivecol():
        def CndShow():

            def addcol():

                def checkcnd():
                    bindx=b1.grid_info()['row']
                    coltoshow=Columnsss[bindx-1]
                    colcndshow=custcolcnd[coltoshow]
                    
                    for widget in frame23.winfo_children():
                        widget.destroy()
                    Cndtext=tk.Text(frame23,bg="white",height = 10, width = 50)
                    Cndtext.grid(row=5,column=1,padx=5,pady=5)
                    Cndtext.delete('1.0',tk.END)
                    Cndtext.insert(tk.END,colcndshow)

                custcolcnd[Colname]=fincnd
                Columnsss=[]
                for i in var1:
                    d=i.cget('text')
                    Columnsss.append(d)
                y=len(Columnsss)
                if Colname not in Columnsss:
                    l2=tk.Label(frame13,text=Colname)
                    l2.grid(row=y+1,column=1,padx=5,pady=3)
                    var1.append(l2)
                    e2=tk.Entry(frame13,width=5)
                    e2.grid(row=y+1,column=2,padx=5,pady=3)
                    var2.append(e2)
                    e3=tk.Entry(frame13,width=10)
                    e3.grid(row=y+1,column=3,padx=5,pady=3)
                    var3.append(e3)
                    b1=tk.Button(frame13,text='Check Condition',command=checkcnd)
                    b1.grid(row=y+1,column=4,padx=5,pady=3)
                    Columnsss.append(Colname)
                else:
                    y=Columnsss.index(Colname) + 1
                    #b2.grid_forget()
                    b1=tk.Button(frame13,text='Check Condition',command=checkcnd)
                    b1.grid(row=y,column=4,padx=5,pady=3)
            
            Colname=combox.get()
            Cndcol=Cndbox.get()
            Opr=Opname.get()
            CndVal=CndEntry.get()        
            Result=Resbox.get()
            addcase=frame27.grid_slaves()
            addcase=addcase[::-1]
            newcnd=[]
            optionwdgt=[]
            if len(addcase)==0:
                fincnd='CASE\n' + 'When ' + Cndcol +' ' + Opr + ' ' + CndVal + '\nThen ' + Result + '\n' + 'END as ' + Colname
            else:
                for widgt in addcase:
                    if isinstance(widgt,tk.OptionMenu):
                        optionwdgt.append(widgt)
                        val=variablelist[optionwdgt.index(widgt)].get()
                    else:
                        val=widgt.get()
                    newcnd.append(val)
                if len(Cndcol)==0:
                    fincnd='CASE\n' + ' '.join(newcnd) + '\n' + 'END as ' + Colname
                else:
                    fincnd='CASE\n' + 'When ' + Cndcol +' ' + Opr + ' ' + CndVal + '\nThen ' + Result + '\n' \
                        + ' '.join(newcnd) + '\n' + 'END as ' + Colname
            Cndtext=tk.Text(frame23,bg="white",height = 10, width = 50)
            Cndtext.grid(row=5,column=1,padx=5,pady=5)
            Cndtext.insert(tk.END,fincnd)
            addbutn=tk.Button(frame24,text='Add',command=addcol)
            addbutn.grid(row=1,column=1,padx=5,pady=5)
            addbutn=tk.Button(frame24,text='Clear',command=addcol)
            addbutn.grid(row=2,column=1,padx=5,pady=5)
        
        for widget in frame14.winfo_children():
            widget.destroy() 
        container = ttk.Frame(frame14,borderwidth=5,relief="solid")
        frame17=tk.Frame(container)
        frame17.grid(sticky=tk.E+tk.W,padx=10, pady=10)
        frameCanvas = tk.Canvas(frame17,width=650,height=350)
        vscrollbar = ttk.Scrollbar(frame17, orient="vertical", command=frameCanvas.yview)
        scrollable_frame = ttk.Frame(frameCanvas)
        scrollable_frame.bind(
                "<Configure>",
                lambda e: frameCanvas.configure(
                scrollregion=frameCanvas.bbox("all")
              )
            )
        frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
        frameCanvas.configure(yscrollcommand=vscrollbar.set)   
        container.grid(row=1,column=5,padx=5,pady=5)
        frameCanvas.pack(side="left")
        vscrollbar.pack(side="right", fill="y")

        frame21=tk.Frame(scrollable_frame)
        frame15=tk.Frame(scrollable_frame)
        frame16=tk.Frame(scrollable_frame)
        frame27=tk.Frame(scrollable_frame)
        frame22=tk.Frame(scrollable_frame)
        frame21.pack()
        frame15.pack()
        frame27.pack()
        frame22.pack()
        frame16.pack()
        frame23=tk.Frame(frame16)
        frame24=tk.Frame(frame16)
        frame23.pack(side='left')
        frame24.pack(side='left')


        def rowcolmax():
            maxcol,maxrow=frame27.grid_size()
            if maxrow!=0:
                maxrow=maxrow-1
            for i in range(5):
                x=len(frame27.grid_slaves(row=maxrow,column=i))
                if x==0:
                    maxcol=i                   
                    break                
            if maxcol==5:
                maxrow +=2
                maxcol=0
            return maxrow,maxcol

        def createDBField():
            mrow,mcol=rowcolmax()
            Equival = tk.StringVar()
            CndEntry=ttk.Combobox(frame27, width = 20, textvariable = Equival,values=availcols)
            
            CndEntry.grid(row=mrow,column=mcol,padx=5,pady=5)
            CndEntry.current()
            variablelist.append(Equival)

        def Opera():
            mrow,mcol=rowcolmax()
            Opname = tk.StringVar()
            Opbox=ttk.Combobox(frame27, width = 5, textvariable = Opname,values=Operatorlist)
            Opbox.current()
            Opbox.grid(row=mrow,column=mcol,padx=5,pady=5)

        def Labels():
            mrow,mcol=rowcolmax()
            widg=frame27.grid_slaves()
            i=0
            
            for x in widg:
                if isinstance(x,tk.OptionMenu):
                    i+=1   
            labli=tk.StringVar()
            Caselabel=['When','Then','is','as','End','if','Select','From','Where','Order by','Group by']
            dropMenu = tk.OptionMenu( frame27 , labli , *Caselabel)
            dropMenu.grid(row=mrow,column=mcol,padx=5,pady=5)
            variablelist.append(labli)           

        def Fnctions():
            mrow,mcol=rowcolmax()
            widg=frame27.grid_slaves()
            i=0
            
            for x in widg:
                if isinstance(x,tk.OptionMenu):
                    i+=1   
            funci=tk.StringVar()
            Fnctions=['Ltrim','Rtrim','Date','DateAdd','Max','Min','Count','Avg','Sum']
            FncMenu = tk.OptionMenu( frame27 , funci , *Fnctions)
            FncMenu.grid(row=mrow,column=mcol,padx=5,pady=5)
            variablelist.append(funci)            

        def remove():
            mrow=rowcolmax()[0]
            widgt=frame27.grid_slaves(row=mrow)
            if len(widgt)!=0:
                widgt[0].destroy()

        

        showbutn=tk.Button(frame21,text='Labels',command=Labels)
        showbutn.grid(row=0,column=0,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='Operator',command=Opera)
        showbutn.grid(row=0,column=1,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='DB field',command=createDBField)
        showbutn.grid(row=0,column=2,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='Functions',command=Fnctions)
        showbutn.grid(row=0,column=3,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='Remove',command=remove)
        showbutn.grid(row=0,column=4,padx=5,pady=5)


        seleclabel=tk.Label(frame15,text='Calculate')
        seleclabel.grid(row=1,column=1,ipady=2,padx=5)
        availcols = listbox2.get(0,tk.END)
        initcname = tk.StringVar()
        combox=ttk.Combobox(frame15, width = 20, textvariable = initcname,values=availcols)
        combox.current()
        combox.grid(row=1,column=2,padx=5,pady=5)
        aslabel=tk.Label(frame15,text='as')
        aslabel.grid(row=1,column=3,ipady=2,padx=5)
        Whenlabel=tk.Label(frame15,text='When')
        Whenlabel.grid(row=2,column=1,ipady=2,padx=5)
        CndCname = tk.StringVar()
        Cndbox=ttk.Combobox(frame15, width = 20, textvariable = CndCname,values=availcols)
        Cndbox.current()
        Cndbox.grid(row=2,column=2,padx=5,pady=5)
        islabel=tk.Label(frame15,text='is')
        islabel.grid(row=2,column=3,ipady=2,padx=5)
        Operatorlist=['=','!=','<','>','>=','<=','in','not in','And','Or','(',')','+','-','*','/']
        Opname = tk.StringVar()
        Opbox=ttk.Combobox(frame15, width = 5, textvariable = Opname,values=Operatorlist)
        Opbox.current()
        Opbox.grid(row=2,column=4,padx=5,pady=5)
        Equival = tk.StringVar()
        CndEntry=ttk.Combobox(frame15, width = 20, textvariable = Equival,values=availcols)
        CndEntry.grid(row=2,column=5,padx=5,pady=5)
        CndEntry.current()
        Thenlabel=tk.Label(frame15,text='Then')
        Thenlabel.grid(row=3,column=1,ipady=2,padx=5)
        ResultVal = tk.StringVar()
        Resbox=ttk.Combobox(frame15, width = 20, textvariable = ResultVal,values=availcols)
        Resbox.current()
        Resbox.grid(row=3,column=2,padx=5,pady=5)
        showbutn=tk.Button(frame22,text='Show',command=CndShow)
        showbutn.grid(row=4,column=1,padx=5,pady=5)

    def finalsql():
        global sql
        Columnsss=[]
        for i in var1:
            d=i.cget('text')
            Columnsss.append(d)
        sequence=[]
        for i in var2:
            d=i.get()
            sequence.append(d)
        alias=[]
        for i in var3:
            d=i.get()
            alias.append(d)
        for seq in sequence:
            finalcolcnd[seq]=[Columnsss[sequence.index(seq)]]
            finalcolcnd[seq].append(alias[sequence.index(seq)])
        for key,value in finalcolcnd.items():
            for x,y in custcolcnd.items():
                if value[0]==x:
                    finalcolcnd[key].append(y)
                else:
                    finalcolcnd[key].append('')
        colsql=dict(sorted(finalcolcnd.items()))
        col=[]
        for key,value in colsql.items():
            if value[2]=='':
                if value[1]=='':
                    col.append('\n' + value[0])
                else:
                    col.append('\n' + value[0] + ' as ' + value[1])
            else:
                col.append('\n' + value[2])
        selcol='Select' + ','.join(col) + '\nFrom'
        sql=re.sub('Select.*from',selcol,EData,flags=re.DOTALL)
        showsql(sql)

    lnddf = pd.read_sql_query(EData,conn)
    custcol=sorted(list(lnddf.columns),key=str.lower)
    custcolcnd={}
    finalcolcnd={}
    #Newreport={}
    
    newWindow = tk.Toplevel(window)
    newWindow.resizable(True, True)
    
    frame8 = tk.Frame(newWindow)
    frame8.config(bd=1, relief=tk.SOLID)
    #frame9 = tk.Frame(newWindow)
    #frame9.config(bd=1, relief=tk.SOLID)
    #frame10=tk.Frame(newWindow)
    
    container = ttk.Frame(frame8)
    frame18=tk.Frame(container)
    frame18.grid(sticky=tk.E+tk.W,padx=10, pady=10)
    frameCanvas = tk.Canvas(frame18,width=1100,height=450)
    
    scrollbar = ttk.Scrollbar(frame18, orient="vertical", command=frameCanvas.yview)
    scrollable_frame = ttk.Frame(frameCanvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: frameCanvas.configure(
        scrollregion=frameCanvas.bbox("all")
      )
    )
    frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
    frameCanvas.configure(yscrollcommand=scrollbar.set)
    
    #frame8.grid(sticky=tk.E+tk.W,padx=10, pady=10)
    frame8.pack(expand=True)
    container.pack(fill="x")
    frameCanvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")


    frame11=tk.Frame(scrollable_frame)
    frame12=tk.Frame(scrollable_frame)
    frame11.grid()
    frame12.grid(sticky=tk.E+tk.W,padx=10, pady=10)
    l1=tk.Label(frame11,text='Report Name')
    l1.grid(row=0,column=2,ipady=2,padx=5)
    global NameE
    NameE=tk.Entry(frame11,width=30)
    NameE.grid(row=0,column=3,ipady=2,padx=5)
    button1=tk.Button(frame11,text='ShowSQL',command=finalsql)
    button1.grid(row=0,column=4,ipady=2,padx=5)
    button2=tk.Button(frame11,text='New Column',command=derivecol)
    button2.grid(row=0,column=5,ipady=2,padx=5)
    #button3=tk.Button(frame11,text='Save Report',command=savereport)
    #button3.grid(row=0,column=6,ipady=2,padx=5)
    button4=tk.Button(frame11,text='Close',command=newWindow.destroy)
    button4.grid(row=0,column=7,ipady=2,padx=5)

    frame13=tk.Frame(frame12)
    frame13.pack(side='left')
    frame14=tk.Frame(frame12)
    frame14.pack(side='right')

    l2=tk.Label(frame13,text='Col Position')
    l2.grid(row=0,column=2,ipady=2)
    l3=tk.Label(frame13,text='Alias')
    l3.grid(row=0,column=3,ipady=2)
    l4=tk.Label(frame13,text='Condition')
    l4.grid(row=0,column=4,ipady=2)
    
    
    y=1
    global var,var1,var2,var3
    var= []
    var1=[]
    var2=[]
    var3=[]
    for x in custcol:
        l2=tk.Label(frame13,text=x)
        l2.grid(row=custcol.index(x)+1,column=y,padx=5,pady=3)
        var1.append(l2)
        e2=tk.Entry(frame13,width=5)
        e2.grid(row=custcol.index(x)+1,column=y+1,padx=5,pady=3)
        var2.append(e2)
        e3=tk.Entry(frame13,width=10)
        e3.grid(row=custcol.index(x)+1,column=y+2,padx=5,pady=3)
        var3.append(e3)
        b2=tk.Button(frame13,text='Add Condition',command=derivecol)
        b2.grid(row=custcol.index(x)+1,column=y+3,padx=5,pady=3)
    
    Infotext=tk.Text(frame14,height =10, width = 80,bg="light Grey")
    Infotext.grid(row=1,column=0,padx=5,pady=5)
    Info='\nReport Name: Specify Name of the report\n\n' \
          'Show Sql: Review the SQL Query\n\n' \
          'DerivedCol: Add and Define New Column\n\n' \
          'Add Condition: Add New Condition for Selected Column\n' \
          'Note: Once a Condition is added to a column, "Add" is replaced by \n "Check Condtion" Button\n\n'
    Infotext.insert(tk.END,Info)

def Customizecolumns(*args):
    def derivecol():
        def CndShow():

            def clearwidgt():
                for widgt in variablelist:
                    widgt.set('')
                for widgt in addcase:
                    if isinstance(widgt,ttk.Combobox):
                        widgt.set('')
                for widgt in frame15.grid_slaves():
                    widgt.set('')

            def addcol():
                newcnd=list(fincnd.split(" "))
                tabwiscnd=[]
                for val in newcnd:
                    if (val.find('(') != -1):
                        fiel=val[:(val.find('('))]
                        fielnmbr=int(val[(val.find('('))+1:(val.find('('))+2])
                        key=list(DB_Table)[fielnmbr-1]
                        value=DB_Table[key][0]
                        datf=value + '.' + key + '.' + fiel
                    else:
                        for key,value in DB_Table.items():
                            if val in value[1]:
                                datf=value[0] + '.' +  key + '.' + val
                            else:
                                datf=val
                    tabwiscnd.append(datf)
            
                if len(combox.get())==0:
                    messagebox.showinfo("Warning","Please specify a Column Name")
                elif combox.get() in globaldict.keys():
                    messagebox.showinfo("Warning","Column Name already exists \n Please specify a different Name")
                else:
                    sqlcnd=' '.join(tabwiscnd)
                    globaldict[combox.get()]=sqlcnd
                    existingderivedcol.append(Colname)
                    Masterdict={'Name':[Colname],'SQL':[sqlcnd],'Type':['Col']}
                    Masterdf=pd.DataFrame(Masterdict)
                    if os.path.isfile(customfilepath):
                        Masterdf.to_csv(customfilepath, mode='a', header=not os.path.exists(customfilepath),index=False)
                    else:
                        Masterdf.to_csv(customfilepath,index=False)
                
            
            Colname=combox.get()
            Cndcol=Cndbox.get()
            Opr=Opname.get()
            CndVal=CndEntry.get()        
            Result=Resbox.get()
            addcase=frame27.grid_slaves()
            addcase=addcase[::-1]
            newcnd=[]
            optionwdgt=[]
            if len(addcase)==0:
                fincnd='CASE ' + '\nWhen ' + Cndcol +' ' + Opr + ' ' + CndVal + '\nThen ' + Result + '\n' + 'END as ' + Colname
            else:
                for widgt in addcase:
                    if isinstance(widgt,tk.OptionMenu):
                        optionwdgt.append(widgt)
                        val=variablelist[optionwdgt.index(widgt)].get()
                    else:
                        val=widgt.get()
                    newcnd.append(val)

                if len(Cndcol)==0:
                    fincnd=' '.join(newcnd)
                else:
                    fincnd='CASE\n' + 'When ' + Cndcol +' ' + Opr + ' ' + CndVal + '\nThen ' + Result + '\n' \
                        + ' '.join(newcnd) + '\n' + 'END as ' + Colname

            
            Cndtext=tk.Text(frame23,bg="white",height = 10, width = 50)
            Cndtext.grid(row=5,column=1,padx=5,pady=5)
            Cndtext.insert(tk.END,fincnd)
            addbutn=tk.Button(frame24,text='Add',command=addcol)
            addbutn.grid(row=1,column=1,padx=5,pady=5)
            addbutn=tk.Button(frame24,text='Clear',command=clearwidgt)
            addbutn.grid(row=2,column=1,padx=5,pady=5)
        
        for widget in frame14.winfo_children():
            widget.destroy() 
        container = ttk.Frame(frame14,borderwidth=5,relief="solid")
        frame17=tk.Frame(container)
        frame17.grid(sticky=tk.E+tk.W,padx=10, pady=10)
        frameCanvas = tk.Canvas(frame17,width=650,height=350)
        vscrollbar = ttk.Scrollbar(frame17, orient="vertical", command=frameCanvas.yview)
        scrollable_frame = ttk.Frame(frameCanvas)
        scrollable_frame.bind(
                "<Configure>",
                lambda e: frameCanvas.configure(
                scrollregion=frameCanvas.bbox("all")
              )
            )
        frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
        frameCanvas.configure(yscrollcommand=vscrollbar.set)   
        container.grid(row=1,column=5,padx=5,pady=5)
        frameCanvas.pack(side="left")
        vscrollbar.pack(side="right", fill="y")

        frame21=tk.Frame(scrollable_frame)
        frame15=tk.Frame(scrollable_frame)
        frame16=tk.Frame(scrollable_frame)
        frame27=tk.Frame(scrollable_frame)
        frame22=tk.Frame(scrollable_frame)
        frame21.pack()
        frame15.pack()
        frame27.pack()
        frame22.pack()
        frame16.pack()
        frame23=tk.Frame(frame16)
        frame24=tk.Frame(frame16)
        frame23.pack(side='left')
        frame24.pack(side='left')


        def rowcolmax():
            maxcol,maxrow=frame27.grid_size()
            if maxrow!=0:
                maxrow=maxrow-1
            for i in range(5):
                x=len(frame27.grid_slaves(row=maxrow,column=i))
                if x==0:
                    maxcol=i                   
                    break                
            if maxcol==5:
                maxrow +=2
                maxcol=0
            return maxrow,maxcol


        def createDBField():
            mrow,mcol=rowcolmax()
            field = tk.StringVar()
            CndEntry=ttk.Combobox(frame27, width = 20, textvariable = field,values=availcols)
            
            CndEntry.grid(row=mrow,column=mcol,padx=5,pady=5)
            CndEntry.current()

        def Opera():
            mrow,mcol=rowcolmax()
            Opname = tk.StringVar()
            Opbox=ttk.Combobox(frame27, width = 5, textvariable = Opname,values=Operatorlist)
            Opbox.current()
            Opbox.grid(row=mrow,column=mcol,padx=5,pady=5)

        
        def Labels():
            mrow,mcol=rowcolmax()
            widg=frame27.grid_slaves()
            i=0
            
            for x in widg:
                if isinstance(x,tk.OptionMenu):
                    i+=1   
            clickedi=tk.StringVar()
            Caselabel=['When','Then','is','as','End','if','Select','From','Where','Order by','Group by']
            dropMenu = tk.OptionMenu( frame27 , clickedi , *Caselabel)
            dropMenu.grid(row=mrow,column=mcol,padx=5,pady=5)
            variablelist.append(clickedi)          

        def Fnctions():
            mrow,mcol=rowcolmax()
            widg=frame27.grid_slaves()
            i=0
            
            for x in widg:
                if isinstance(x,tk.OptionMenu):
                    i+=1   
            clickedi=tk.StringVar()
            Fnctions=['Ltrim','Rtrim','Date','DateAdd','Max','Min','Count','Avg','Sum']
            FncMenu = tk.OptionMenu( frame27 , clickedi , *Fnctions)
            FncMenu.grid(row=mrow,column=mcol,padx=5,pady=5)
            variablelist.append(clickedi)        

        def remove():
            addcase=frame27.grid_slaves()
            addcase=addcase[::-1]
            if len(addcase)!=0:
                if isinstance(addcase[-1],tk.OptionMenu):
                    variablelist.pop()
                addcase[-1].destroy()
                addcase.pop()
                
        

        showbutn=tk.Button(frame21,text='Labels',command=Labels)
        showbutn.grid(row=0,column=0,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='Operator',command=Opera)
        showbutn.grid(row=0,column=1,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='DB field',command=createDBField)
        showbutn.grid(row=0,column=2,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='Functions',command=Fnctions)
        showbutn.grid(row=0,column=3,padx=5,pady=5)
        showbutn=tk.Button(frame21,text='Remove',command=remove)
        showbutn.grid(row=0,column=4,padx=5,pady=5)


        seleclabel=tk.Label(frame15,text='Calculate')
        seleclabel.grid(row=1,column=1,ipady=2,padx=5)
        initcname = tk.StringVar()
        combox=ttk.Combobox(frame15, width = 20, textvariable = initcname,values=availcols)
        combox.current()
        combox.grid(row=1,column=2,padx=5,pady=5)
        aslabel=tk.Label(frame15,text='as')
        aslabel.grid(row=1,column=3,ipady=2,padx=5)
        Whenlabel=tk.Label(frame15,text='When')
        Whenlabel.grid(row=2,column=1,ipady=2,padx=5)
        CndCname = tk.StringVar()
        Cndbox=ttk.Combobox(frame15, width = 20, textvariable = CndCname,values=availcols)
        Cndbox.current()
        Cndbox.grid(row=2,column=2,padx=5,pady=5)
        islabel=tk.Label(frame15,text='is')
        islabel.grid(row=2,column=3,ipady=2,padx=5)
        Opname = tk.StringVar()
        Opbox=ttk.Combobox(frame15, width = 5, textvariable = Opname,values=Operatorlist)
        Opbox.current()
        Opbox.grid(row=2,column=4,padx=5,pady=5)
        Equival = tk.StringVar()
        CndEntry=ttk.Combobox(frame15, width = 20, textvariable = Equival,values=availcols)
        CndEntry.grid(row=2,column=5,padx=5,pady=5)
        CndEntry.current()
        Thenlabel=tk.Label(frame15,text='Then')
        Thenlabel.grid(row=3,column=1,ipady=2,padx=5)
        ResultVal = tk.StringVar()
        Resbox=ttk.Combobox(frame15, width = 20, textvariable = ResultVal,values=availcols)
        Resbox.current()
        Resbox.grid(row=3,column=2,padx=5,pady=5)
        showbutn=tk.Button(frame22,text='Show',command=CndShow)
        showbutn.grid(row=4,column=1,padx=5,pady=5)

    variablelist=[]
        
        

    newWindow = tk.Toplevel(window)
    newWindow.resizable(True, True)
    
    frame8 = tk.Frame(newWindow)
    frame8.config(bd=1, relief=tk.SOLID)
    #frame9 = tk.Frame(newWindow)
    #frame9.config(bd=1, relief=tk.SOLID)
    #frame10=tk.Frame(newWindow)
    
    container = ttk.Frame(frame8)
    frame18=tk.Frame(container)
    frame18.grid(sticky=tk.E+tk.W,padx=10, pady=10)
    frameCanvas = tk.Canvas(frame18,width=1100,height=450)
    
    scrollbar = ttk.Scrollbar(frame18, orient="vertical", command=frameCanvas.yview)
    scrollable_frame = ttk.Frame(frameCanvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: frameCanvas.configure(
        scrollregion=frameCanvas.bbox("all")
      )
    )
    frameCanvas.create_window((0,0), window=scrollable_frame, anchor="nw")
    frameCanvas.configure(yscrollcommand=scrollbar.set)
    
    #frame8.grid(sticky=tk.E+tk.W,padx=10, pady=10)
    frame8.pack(expand=True)
    container.pack(fill="x")
    frameCanvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    frame11=tk.Frame(scrollable_frame)
    frame12=tk.Frame(scrollable_frame)
    frame11.grid()
    frame12.grid(sticky=tk.E+tk.W,padx=10, pady=10)
    button2=tk.Button(frame11,text='New Column',command=derivecol)
    button2.grid(row=0,column=5,ipady=2,padx=5)
    #button3=tk.Button(frame11,text='Save Report',command=savereport)
    #button3.grid(row=0,column=6,ipady=2,padx=5)
    button4=tk.Button(frame11,text='Close',command=newWindow.destroy)
    button4.grid(row=0,column=7,ipady=2,padx=5)

    frame13=tk.Frame(frame12)
    frame13.pack(side='left')
    frame14=tk.Frame(frame12)
    frame14.pack(side='right')

def createreport(*args):
    tablename = [listbox.get(i) for i in listbox.curselection()]
    Reportname = [listbox3.get(i) for i in listbox3.curselection()]
    ExceptionRpt=ERpt.get()
    MultiRpt=MRpt.get()
    XcelRpt=XRpt.get()
    CSVRpt=CRpt.get()
    EReport={}
    colRpt={}
    rowRpt={}
    tabcoldict={}
    CntReport={}
    # if no table or report selected
    if len(Reportname)==0 and len(tablename)==0:
        cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=? and table_name like 'stg_%'",dbname)
        all_tables=[item[0] for item in cursor.fetchall()]
        for val in all_tables:
            EData="Select * from {} where Recordstatus='E'".format(val)
            key=val.split("_",1)[1] + "_Exception"
            lnddf = pd.read_sql_query(EData,conn)
            rowRpt[key] = lnddf.shape[0]  # Gives number of rows
            colRpt[key] = lnddf.shape[1]  # Gives number of col
            EReport[key]=lnddf
            Errordf=lnddf.groupby(["ErrorReason"])["ErrorReason"].value_counts()
            CntReport[key]=Errordf
        print("general file")
    # if table selected but no report
    if len(tablename)!=0 and len(Reportname)==0:
    #if specific columns required
        if frame11.winfo_children():
            for i in range(len(colvar)):
                tabcoldict[tablename[i]]=colvar[i].get(0,tk.END)
            for key,value in tabcoldict.items():
                if len(value)==0:
                    value='*'
                selcol=','.join(value)
                if ExceptionRpt==1:
                    EData="Select " + selcol + ' from ' + key + " where Recordstatus='E'"
                    key= key.split("_",1)[1] + "_Exception"
                else:
                    EData="Select " + selcol + ' from ' + key
                lnddf = pd.read_sql_query(EData,conn)
                rowRpt[key] = lnddf.shape[0]  # Gives number of rows
                colRpt[key] = lnddf.shape[1]  # Gives number of col
                EReport[key]=lnddf
                if ExceptionRpt==1:
                    Errordf=lnddf.groupby(["ErrorReason"])["ErrorReason"].value_counts()
                    CntReport[key]=Errordf
            print("col specific file")
        # if only tables selected
        else: 
            for val in tablename:
                if ExceptionRpt==1:
                    EData="Select * from {} where Recordstatus='E'".format(val)
                    val= val.split("_",1)[1] + "_Exception"
                else:
                    EData="Select * from " +  val
                lnddf = pd.read_sql_query(EData,conn)
                rowRpt[val] = lnddf.shape[0]  # Gives number of rows
                colRpt[val] = lnddf.shape[1]  # Gives number of col
                EReport[val]=lnddf
                if ExceptionRpt==1:
                    Errordf=lnddf.groupby(["ErrorReason"])["ErrorReason"].value_counts()
                    CntReport[val]=Errordf
            print("table specific file")
    #code for only report name
    if len(tablename)==0 and len(Reportname)!=0:
        for val in Reportname:
            if 'Exception' in val:
               key = "Stg_" + val.split("_",1)[0]
               EData="Select * from {} where Recordstatus='E'".format(key)
            lnddf = pd.read_sql_query(EData,conn)
            rowRpt[val] = lnddf.shape[0]  # Gives number of rows
            colRpt[val] = lnddf.shape[1]  # Gives number of col
            EReport[val]=lnddf
            #Errordf=lnddf.groupby(["ErrorReason"])["ErrorReason"].value_counts()
            Errordf=lnddf.groupby(["ErrorReason"]).size()
            CntReport[key]=Errordf
        print(CntReport)   
        print("report specific file")
    #for future use to give user option to name file and location
    #export_file_path = filedialog.asksaveasfilename(defaultextension='.csv') 
     
    #Final Rpt Creation
    if XcelRpt==1:
        if MultiRpt==0: 
            writer = ExcelWriter('Exception.xlsx') # pylint: disable=abstract-class-instantiated
            for key in EReport:
                EReport[key].to_excel(writer, key,index=False,engine='xlsxwriter',startrow=20)
                #CntReport[key].to_excel(writer, key,index=False,engine='xlsxwriter',startrow=3,startcol=2)
        else:
            for key in EReport:
                Rname=key +'.xlsx'
                writer = ExcelWriter(Rname) # pylint: disable=abstract-class-instantiated
                EReport[key].to_excel(writer, key,index=False,engine='xlsxwriter',startrow=20)
                #CntReport[key].to_excel(writer, key,index=False,engine='xlsxwriter',startrow=3,startcol=2)
                workbook = writer.book # pylint: disable=no-member
                worksheet = writer.sheets[key]
                worksheet.set_row(0,25)
                worksheet.set_column(0,colRpt[key],20)


                blank_fmt = workbook.add_format({'bg_color': '#FFFFFF','border': 1,'border_color':'#d3d3d3'}) # light gray border color
                worksheet.conditional_format(21,colRpt[key],rowRpt[key]+21,16334, {'type': 'blanks',
                                     'format': blank_fmt})
                # Add Row color
                bg_format1 = workbook.add_format({'bg_color': '#E0FFFF','border': 0}) # light cyan cell background color
                bg_format2 = workbook.add_format({'bg_color': '#AFEEEE','border': 0}) # pale turquoise background color

                for i in range(rowRpt[key]): # integer odd-even alternation 
                    worksheet.set_row(i+21, cell_format=(bg_format1 if i%2==0 else bg_format2))
                print(colRpt[key])
                
                #set zoom
                worksheet.set_zoom(80)
                # Add a header format
                header_fmt = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'bg_color': '#D7E4BC',
                    'border': 1})
                worksheet.conditional_format(20,0,0,colRpt[key], {'type': 'no_blanks',
                                     'format': header_fmt})
                # Add column filter
                worksheet.autofilter(20, 0,rowRpt[key],colRpt[key] - 1)

                tabborder=workbook.add_format({'border':1})
                worksheet.conditional_format(20,0,rowRpt[key]+21,colRpt[key]-1, {'type': 'no_errors',
                                       'format': tabborder})

                writer.save()
                writer.close()
    if CSVRpt==1:
        for key in EReport:
            Rname=key +'.csv'
            EReport[key].to_csv(Rname, index = False, header=True) 

def clearall():
    window.destroy()      
        
EData=''
Wh=''
Jtype=''
Cnd='' 
DB_Table={}
globaldict={}
availcols=[]
choices = {}
servername=''
dbname=[]
variablelist=[]
selcolms=[]
Operatorlist=['=','!=','<','>','>=','<=','in','not in','And','Or','(',')','+','-','*','/']

conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                  Server='localhost',
                  Database='Master',
                  Trusted_Connection='yes',autocommit=True)
cursor = conn.cursor()
Query='Select name from sys.databases'
cursor.execute(Query)
Databases=sorted([item[0] for item in cursor.fetchall()],key=str.lower)

currdir=os.getcwd()
customfilepath=currdir + '/' + 'CustomReportandColumns.csv'
isFile = os.path.isfile(customfilepath)
existingderivedcol=[]
existingreport=[]
if isFile==True:
    CRCDf = pd.read_csv(customfilepath,header=0)
    globaldict=CRCDf.set_index('Name').T.to_dict('list')
    for key,value in globaldict.items():
        if value[1]=='Rpt':
            existingreport.append(key)
        if value[1]=='Col':
            existingderivedcol.append(key)



label1=tk.Label(frame2,text='Source Table',bg="white")
label1.grid(row=2)
listbox =tk.Listbox(frame2,height=5,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",exportselection=False,selectmode = "multiple")
listbox.grid(row=2,column=2)
scrollbar = tk.Scrollbar(frame2)
scrollbar.grid(row = 2,rowspan = 1, column = 3,ipady=15)
listbox.config(yscrollcommand = scrollbar.set)
scrollbar.config(command = listbox.yview)

label2=tk.Label(frame2,text='Reports',bg="white")
label2.grid(row=2,column=10,padx=10)
listbox3 =tk.Listbox(frame2,height=5,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",exportselection=False,selectmode = "multiple")
listbox3.grid(row=2,column=11)
scrollbar3 = tk.Scrollbar(frame2)
scrollbar3.grid(row = 2,rowspan = 1, column = 12,ipady=15)
listbox3.config(yscrollcommand = scrollbar3.set)
scrollbar3.config(command = listbox3.yview)
button2=tk.Button(frame2,text='...',command=reports)
button2.grid(row=2,column=13)

label3=tk.Label(frame2,text='Custom Columns',bg="white")
label3.grid(row=2,column=6,padx=10)
listbox4 =DragDropListbox(frame2,height=5,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",exportselection=False,selectmode = "multiple")
listbox4.grid(row=2,column=7)
scrollbar4 = tk.Scrollbar(frame2)
scrollbar4.grid(row = 2,rowspan = 1, column = 8,ipady=15)
listbox4.config(yscrollcommand = scrollbar4.set)
scrollbar4.config(command = listbox4.yview)
button2=tk.Button(frame2,text='...',command=reports)
button2.grid(row=2,column=9)

optionframe=tk.Frame(frame13)
optionframe.grid(row=4,column=1)
FTframe=tk.Frame(optionframe)
FTframe.pack(padx=5,pady=3)
XRpt = tk.IntVar()
chbox1=tk.Checkbutton(FTframe,text='XLSX',variable=XRpt)
chbox1.pack(side='left',padx=5,pady=3)
CRpt = tk.IntVar()
chbox2=tk.Checkbutton(FTframe,text='CSV',variable=CRpt)
chbox2.pack(side='left',padx=5,pady=3)
ERpt = tk.IntVar()
chbox3=tk.Checkbutton(optionframe,text='Exception Report',variable=ERpt)
chbox3.pack(padx=5,pady=3)
MRpt = tk.IntVar()
chbox4=tk.Checkbutton(optionframe,text='Separate RPT for each table',variable=MRpt)
chbox4.pack(padx=5,pady=3)




servdetailbut= tk.Button(frame13,text="Server",command=SerDetail)
servdetailbut.grid(row=1,column=1,padx=10, pady=10)

RPTbutton=tk.Button(optionframe,text='Get RPT',command=createreport)
RPTbutton.pack(padx=5,pady=3)

button7=tk.Button(frame13,text='Close',command=clearall)
button7.grid(row=5,column=1,padx=10, pady=10)
button8=tk.Button(frame13,text='Create DBFields',command=Customizecolumns)
button8.grid(row=3,column=1,padx=10, pady=10)
button3=tk.Button(frame2,text='Show Columns',command=getlndcolumns)
button3.grid(row=2,column=5,padx=10, pady=10)

frame13.pack(anchor="ne")
frame9.grid(row=2,column=10,padx=10, pady=10)
frame10.pack(side="left")
frame11.pack(side="left",padx=5)
frame12.pack(side="left",padx=5)
frame6.pack(side="bottom")
frame1.grid(sticky=tk.E+tk.W,padx=10, pady=10)
frame2.grid(sticky=tk.E+tk.W,padx=10, pady=10)
frame3.grid(sticky=tk.E+tk.W)
frame7.grid(sticky=tk.E+tk.W)
frame8.grid()
myCanvas2.pack(expand=True,side='right')
myCanvas.pack(expand=True,side='left')
frame4.grid()
window.mainloop()
