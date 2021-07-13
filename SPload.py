import pyautogui
import tkinter as tk
from tkinter import ttk
from tkinter import font
import pyodbc
import pandas as pd
import re
from datetime import datetime
window = tk.Tk()
#window.geometry("1200x800")
window.state('zoomed')
window.configure(bg="white")
window.title('Welcome')
myCanvas = tk.Canvas(window, bg="white")
myCanvas2 = tk.Canvas(window, bg="white",highlightbackground="black",highlightthickness=1)
frame13 = tk.Frame(myCanvas2,highlightbackground="black",highlightthickness=1)
frame13.pack(anchor="ne")
frame1 = tk.Frame(myCanvas,bg="white")
frame2 = tk.Frame(myCanvas,bg="white")
frame9 = tk.Frame(frame2,bg="white")
frame9.grid(row=2,column=10,padx=10, pady=10)
frame3 = tk.Frame(myCanvas,bg="white")
frame10 = tk.Frame(frame3,bg="white")
frame11 = tk.Frame(frame3,bg="white")
frame12 = tk.Frame(frame3,bg="white")
frame10.pack(side="left")
frame11.pack(side="left")
frame12.pack(side="left")

frame4 = tk.Frame(myCanvas,bg="white")
frame5 = tk.Frame(myCanvas,bg="white")
frame6 = tk.Frame(frame3,bg="white")
frame6.pack(side="bottom")
#frame6.grid(row=3,column=2,pady=5)
frame7 = tk.Frame(myCanvas,bg="white")
frame8 = tk.Frame(myCanvas,bg="white",highlightbackground="black",highlightthickness=1)

pyautogui.click(100, 100)
Heading=pyautogui.typewrite("Staging Layer Load")
label= tk.Label(frame1,text=Heading,justify=tk.CENTER,bg="white").pack()

#%store -r servername 
#%store -r dbname

def SerDetail():
    global Serve2,Serve3,ServDetail
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


    l2=tk.Label(frameCanvas,text='Server Name').grid(row=2,padx=5,pady=5)
    Serve2 = tk.Entry(frameCanvas,width=20)
    Serve2.grid(row=2, column=2,padx=5,pady=5)

    l3=tk.Label(frameCanvas,text='Database Name').grid(row=3,padx=5,pady=5)
    Serve3 = tk.Entry(frameCanvas,width=20)
    Serve3.grid(row=3, column=2,padx=5,pady=5)
    Button= tk.Button(frameCanvas,text="Submit",command=SetServerdet).grid(row=4, column=2,padx=5,pady=5)
    Button1= tk.Button(frameCanvas,text="Close",command=ServDetail.destroy).grid(row=4,padx=5,pady=5)

def SetServerdet():
    
    global servername,dbname,cursor,conn
    if servername=='':
        servername=Serve2.get()
    if dbname=='':
        dbname=Serve3.get()
    ServDetail.destroy()  

    conn = pyodbc.connect(Driver='{SQL Server Native Client 11.0}',
                      Server=servername,
                      Database=dbname,
                      Trusted_Connection='yes',autocommit=True)
    cursor = conn.cursor()
    
    lndcallback()

Button= tk.Button(frame13,text="Server",command=SerDetail).grid(row=1,column=1,padx=10, pady=10)

def lndcallback(): 
    listbox.delete(0,tk.END)
    #cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=? and table_name like '%lnd_%'",dbname)
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=?",dbname)
    lnd_tables=sorted([item[0] for item in cursor.fetchall()],key=str.lower)
    for i in lnd_tables:
        listbox.insert(tk.END,i)
    stgcallback()


label1=tk.Label(frame2,text='Select Source Table',bg="white").grid(row=2)
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
button1=tk.Button(frame2,text='...',command=lndcallback).grid(row=2,column=4)

def getlndcolumns():
    global tablename
    global cols
    global listbox2
    global tabindex
    #tablename=listbox.get(listbox.curselection())

    label2=tk.Label(frame10,text='Source Columns   ',bg="white").grid(row=2,ipadx=10)
    listbox2 =DragDropListbox(frame10,height=15,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black",exportselection=0)
    listbox2.grid(row=2,column=2)
    scrollbar2 = tk.Scrollbar(frame10)
    scrollbar2.grid(row = 2,rowspan = 1, column = 3,ipady=15)
    listbox2.config(yscrollcommand = scrollbar2.set)
    scrollbar2.config(command = listbox2.yview)
    button5=tk.Button(frame10,text='Show data',command=getdata).grid(row=3,column=2,pady=10)
    """
    if len(listbox.curselection())>2:
        listbox.selection_clear(0, tk.END)
        listbox2.insert(tk.END,'/* Select (max) two tables */')
        tablename=''
    else:
        tablename = [listbox.get(i) for i in listbox.curselection()]
    """
    tablename = [listbox.get(i) for i in listbox.curselection()]
    tabindex=[]
    for val in tablename:
        top10lnd='Select top 10 * from {}'.format(val)
        lnddf = pd.read_sql_query(top10lnd, conn)
        cols=sorted(list(lnddf.columns),key=str.lower)
        tabname="----" + val + "----"
        listbox2.insert(tk.END,tabname)
        index = listbox2.get(0, "end").index(tabname)
        tabindex.append(index)
        listbox2.itemconfig(index, bg='yellow')
        for col in cols:
            listbox2.insert(tk.END,col)
            
    global joincondition,jointype,Wherecondition1,Wherecondition2
    l4=tk.Label(frame11,text=' Where Clause',bg="white")
    l4.grid(row=5,column=5,padx=5,pady=3)
    Wherecondition1=tk.Text(frame11,width=25,height=3)
    Wherecondition1.grid(row=5,column=6,padx=5,pady=3,ipady=10)
    sb9 = tk.Scrollbar(frame11, orient=tk.VERTICAL, command=Wherecondition1.yview)
    sb9.grid(row=5,column=7,ipady=3)
    Wherecondition1.config(yscrollcommand=sb9.set)
    
    if len(listbox.curselection())>1:
        join=['Inner','Left','Right','Outer','Union','Union All']
        
        #joinvar = tk.StringVar() 
        #joinvar.set('Select') # default value

        #if len(tablename)>1:
        """
        l2=tk.Label(frame11,text='  Join Type',bg="white")
        l2.grid(row=3,column=5,padx=5,pady=3)
        O2=tk.OptionMenu(frame11,joinvar,*join)
        O2.grid(row=3,column=6,pady=3)
        """
        l4=tk.Label(frame11,text=' Join Type',bg="white")
        l4.grid(row=3,column=5,padx=5,pady=3)
        jointype=tk.Text(frame11,width=25,height=1.5)
        jointype.grid(row=3,column=6,padx=5,pady=3)
        
        l3=tk.Label(frame11,text=' Join Condition',bg="white")
        l3.grid(row=4,column=5,padx=5,pady=3)
        joincondition=tk.Text(frame11,width=25,height=3)
        joincondition.grid(row=4,column=6,padx=5,pady=3,ipady=10)
        sb9 = tk.Scrollbar(frame11, orient=tk.VERTICAL, command=joincondition.yview)
        sb9.grid(row=4,column=7,ipady=3)
        joincondition.config(yscrollcommand=sb9.set)
        
        """
        l4=tk.Label(frame11,text=' Where Clause',bg="white")
        l4.grid(row=5,column=5,padx=5,pady=3)
        Wherecondition1=tk.Text(frame11,width=25,height=3)
        Wherecondition1.grid(row=5,column=6,padx=5,pady=3,ipady=10)
            
        l5=tk.Label(frame11,text=' Table2 Condition',bg="white")
        l5.grid(row=6,column=5,padx=5,pady=3)
        #Wherecondition2=tk.Entry(frame11,width=30)
        #Wherecondition2.grid(row=6,column=6,padx=5,pady=3)
        """         
    
    
    #dnd = DragManager()
    #dnd.add_dragable(listbox2)
  
def getdata(*args):
    #query='Select top 10 * from {}'.format(tablename)
    #df = pd.read_sql_query(query, conn)
    for val in (args or tablename):
        top10lnd='Select top 10 * from {}'.format(val)
        lnddf = pd.read_sql_query(top10lnd, conn)
        lndcols=list(lnddf.columns)
        newWindow = tk.Toplevel(window)
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
        label3=tk.Label(scrollable_frame,text='{} table data for reference'.format(val)).grid(row=0,column=0)
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
        button6=tk.Button(scrollable_frame,text='Close data',command= newWindow.destroy).grid(row=2,column=0)
    
    
    
#def closedata():
    #for widget in frame7.winfo_children():
        #widget.destroy()
    


def stgcallback(): 
    listbox3.delete(0,tk.END)
    #cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=? and table_name like '%stg_%'",dbname)
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=?",dbname)
    stg_tables=sorted([item[0] for item in cursor.fetchall()],key=str.lower)
    for i in stg_tables:
        listbox3.insert(tk.END,i)        
        

label2=tk.Label(frame2,text='Select Target Table',bg="white").grid(row=2,column=6,padx=10)
listbox3 =tk.Listbox(frame2,height=5,  
                  width = 30,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black")
listbox3.grid(row=2,column=7)
scrollbar3 = tk.Scrollbar(frame2)
scrollbar3.grid(row = 2,rowspan = 1, column = 8,ipady=15)
listbox3.config(yscrollcommand = scrollbar3.set)
scrollbar3.config(command = listbox3.yview)
button2=tk.Button(frame2,text='...',command=stgcallback).grid(row=2,column=9)

def getstgcolumns():
    global stgtable
    global stgcols
    global stgdf
    global stgbutton
    for widget in frame12.winfo_children():
            widget.destroy() 
    stgtable=listbox3.get(listbox3.curselection())
        
    label4=tk.Label(frame12,text='Target Columns',bg="white").grid(row=2,column=10,ipadx=20)
    listbox4 =tk.Listbox(frame12,height=15,  
                  width = 25,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black")
    listbox4.grid(row=2,column=11)
    scrollbar4 = tk.Scrollbar(frame12)
    scrollbar4.grid(row = 2,rowspan = 1, column = 12,ipady=15)
    listbox4.config(yscrollcommand = scrollbar4.set)
    scrollbar4.config(command = listbox4.yview)
    stgbutton=tk.Button(frame12,text='Show data',command=getstgdata).grid(row=3,column=11,pady=10)
    query='Select top 10 * from {}'.format(stgtable)
    stgdf = pd.read_sql_query(query, conn)
    stgcols=list(stgdf.columns)
    for col in stgcols:
        listbox4.insert(tk.END,col)

def getstgdata():
    
    query='Select top 10 * from {}'.format(stgtable)
    stgdf = pd.read_sql_query(query, conn)
    cols=list(stgdf.columns)
    
    newWindow = tk.Toplevel(window)
    newWindow.resizable(True, False)
    newCanv = tk.Canvas(newWindow, bg="white")
    newCanv.pack()
    label3=tk.Label(newCanv,text='{} table data for reference'.format(stgtable),bg="white").grid(row=0,column=0)
    tree = ttk.Treeview(newCanv, height=10,columns=cols, show='headings')
    tree.grid(row=1, column=0, sticky='news',padx=10)
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=35, anchor=tk.CENTER)
    for i in range(len(stgdf)):
        tree.insert('','end', values=list(stgdf.loc[i]))
        
    sb = tk.Scrollbar(newCanv, orient=tk.VERTICAL, command=tree.yview)
    sb.grid(row=1, column=1, sticky='ns')
    tree.config(yscrollcommand=sb.set)
    button6=tk.Button(newCanv,text='Close data',command= newWindow.destroy).grid(row=2,column=0)
    

def mapcolumns():
    #frameCanvas = tk.Canvas(frame8, bg="white",height=1000, width=1000)
    #frameCanvas.grid(ipadx=10,ipady=15)
    table=''
    if stgtable!=table:
        for widget in frame8.winfo_children():
            widget.destroy() 
    
        container = ttk.Frame(frame8)
        frameCanvas = tk.Canvas(container, bg="white",height=200, width=1050)
    #frameCanvas.grid(ipadx=10,ipady=15)
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
    
        container.pack(fill="both", expand=True)
        frameCanvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")
    
    
        l1=tk.Label(scrollable_frame,text='Landing Column')
        l1.grid(row=0,column=2,ipady=2)
        l2=tk.Label(scrollable_frame,text='Condition')
        l2.grid(row=0,column=3,ipady=2)
        l3=tk.Label(scrollable_frame,text='Landing Column')
        l3.grid(row=0,column=6,ipady=2)
        l4=tk.Label(scrollable_frame,text='Condition')
        l4.grid(row=0,column=7,ipady=2)
        button8=tk.Button(frame11,text='Show SP',command=showSP).grid(row=7,column=6,padx=10, pady=10)
        y=1
        global var,var1,var2,var3
        var= []
        var1=[]
        var2=[]
        var3=[]
        mI = (len(stgcols)-1)//2 

        for x in stgcols[1:mI+1]:
            l2=tk.Label(scrollable_frame,text=x)
            l2.grid(row=stgcols.index(x),column=y,padx=5,pady=3)
            e2=tk.Entry(scrollable_frame,width=30)
            e2.grid(row=stgcols.index(x),column=y+1,padx=5,pady=3)
            var.append(e2)
            e3=tk.Entry(scrollable_frame,width=30)
            e3.grid(row=stgcols.index(x),column=y+2,padx=5,pady=3)
            var1.append(e3)
        y=5  
        for x in stgcols[mI+1:]:
            l2=tk.Label(scrollable_frame,text=x)
            l2.grid(row=stgcols.index(x)-mI,column=y,padx=5,pady=3)
            e2=tk.Entry(scrollable_frame,width=30)
            e2.grid(row=stgcols.index(x)-mI,column=y+1,padx=5,pady=3)
            var2.append(e2)
            e3=tk.Entry(scrollable_frame,width=30)
            e3.grid(row=stgcols.index(x)-mI,column=y+2,padx=5,pady=3)
            var3.append(e3)
                  
class DragDropListbox(tk.Listbox):
    def __init__(self, master, **kw):
        kw['selectmode'] = tk.SINGLE
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.setCurrent)
        #self.bind('<B1-Motion>', self.shiftSelection)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonRelease-1>", self.on_drop)
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

    def on_leave(self, event):
        global value
        if self.curIndex == None:
            pass
        else:
            selection=self.curselection()
            value=self.get(selection[0])
            #valind=self.get(0, "end").index(value)
            i=self.curIndex
            for x in tabindex:
                if x>i:
                    break
                z=tabindex.index(x)
            value='tab' + str(z+1) + '.' + value
    
    def on_drop(self, event):
        # find the widget under the cursor
        x,y = event.widget.winfo_pointerxy()
        target = event.widget.winfo_containing(x,y)
        try:
            #target.delete(0,tk.END)
            target.insert(tk.END,value)
        except:
            pass
            

def showSP():
    insrtcol= ',\n'.join(stgcols[1:])
    
    selct1=[]
    selct3=[]
    for i in range(len(var)):
        C=var[i].get()
        selct1.append(C)
        D=var1[i].get()
        selct3.append(D)
        
    selct1=["''" if x == '' else x for x in selct1]
    selct3=["" if x == '' else x for x in selct3]
    selct100=[]
    for i in range(len(selct1)):
        if selct3[i]!='':
            selct100.append(selct3[i])
        else:
            selct100.append(selct1[i])
    
    
    selct2=[]
    selct4=[]
    for i in range(len(var2)):
        C=var2[i].get()
        selct2.append(C)
        D=var3[i].get()
        selct4.append(D)
    selct2=["''" if x == '' else x for x in selct2]
    selct4=["" if x == '' else x for x in selct4]
    selct101=[]
    for i in range(len(selct2)):
        if selct4[i]!='':
            selct101.append(selct4[i])
        else:
            selct101.append(selct2[i])
    selct100.extend(selct101)  
        
    selct='\n,'.join([selct100[i]+' AS ' + stgcols[i+1] for i in range(len(selct100))])
    
    #CODE for handling Union
    RepCom=["," if x=="''" else x for x in selct100]
    splitselct=[item for items in RepCom for item in items.split(",")]
    fixselct=["''" if x == '' else x for x in splitselct]
    Uselct1= fixselct[::2]
    Uselct2= fixselct[1::2]
    Unionselect1='\n,'.join([Uselct1[i]+' AS ' + stgcols[i+1] for i in range(len(Uselct1))])
    Unionselect2='\n,'.join([Uselct2[i]+' AS ' + stgcols[i+1] for i in range(len(Uselct2))])
    #-----
    
    tablename = [listbox.get(i) for i in listbox.curselection()]
    Wh=Wherecondition1.get("1.0","end-1c")
    
    if len(Wh)==0:
        Where=''
        Where1=''
        Where2=''
    else:
        W=Wh.replace(',','\nand\n')
        if ';' in W:
            W1=W.split(';')
            Where1='where ' + W1[0]
            Where2='where ' + W1[1]
        else:
            Where='where ' + W
        #W1=W.split(';')[0]
        #W2=W.split(';')[1]
    """ 
        if len(W1)==0:
            Where1=''
        elif ',' in W1:
            Where1='where ' + W1.replace(',','\nand\n')
        else:   
            Where1='where ' + W1
    
        if len(W2)==0:
            Where2=''
        elif ',' in W2:
            Where2='where ' + W2.replace(',','\nand\n')
        else:
            Where2='where ' + W2
       
    if len(W2)==0 and len(W1)!=0:
        Where3=Where1
    elif len(W2)==0 and len(W1)!=0:
        Where3=Where2
    elif len(W2)!=0 and len(W1)!=0:
        Where3=Where1 + '\nand\n' + W2.replace(',','\nand\n')
    else:
        Where3=''
    """
    
    
    if len(tablename)==1:
        insrtSP = 'CREATE PROCEDURE uspIntegrate_{} \nAS '.format(stgtable) + '\nBEGIN ' +'\n   SET NOCOUNT ON;\n'+ ' \nInsert into ' + stgtable + ' \n(\n ' + insrtcol  + ' ) ' \
        ' \n(\n ' + '\nSelect\n'+ selct + '\nfrom\n' + tablename[0]  + '\n' + Where + ' \n) ' +'\nEND'
    else:
        insrt= 'CREATE PROCEDURE uspIntegrate_{} \nAS '.format(stgtable) + '\nBEGIN ' +'\n   SET NOCOUNT ON;\n'+ \
            ' \nInsert into ' + stgtable + ' \n(\n ' + insrtcol  + ' ) '
        
        Jtype=jointype.get("1.0","end-1c")
        Cnd=joincondition.get("1.0","end-1c")
        
        if 'Union' in Jtype:
##############Union Works for two tables only##############
            insrtSP = insrt + \
            ' \n(\n ' + '\nSelect\n'+ Unionselect1 + '\nfrom\n' + tablename[0]  + ' tab1\n' + Where1 + ' \n) ' + Jtype + \
            ' \n(\n ' + '\nSelect\n'+ Unionselect2 + '\nfrom\n' + tablename[1]  + ' tab2\n' + Where2 + ' \n) '
            '\nEND'
        else:
            
            Jtype=Jtype.split(',')
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
                
            insrtSP = insrt + ' \n(\n ' + '\nSelect\n'+ selct + '\nfrom\n' \
            + finalfrmcmd + Where + ' \n) ' +'\nEND'
            #tablename[0]  + ' tab1\n'  + Jtype + ' join ' + tablename[1] + ' tab2\nON\n' + JCondition 
            
        Jtype=''      
    
    
    newWindow = tk.Toplevel(window)
    newCanv = tk.Canvas(newWindow, bg="white")
    newCanv.pack(expand=True)
    label3=tk.Label(newCanv,text='uspIntegrate_{}'.format(stgtable),bg="white").grid(row=0,column=0)
    global textbox
    textbox=tk.Text(newCanv,bg="white",height = 20, width = 100)
    textbox.grid(row=1,column=0,padx=50,pady=50)
    
    sb = tk.Scrollbar(newCanv, orient=tk.VERTICAL, command=textbox.yview)
    sb.grid(row=1,column=1,ipady=50)
    textbox.config(yscrollcommand=sb.set)
    button1=tk.Button(newCanv,text='Save/Execute SP',command= save_execute_SP).grid(row=2,column=0,ipady=5)
    button2=tk.Button(newCanv,text='Close',command= newWindow.destroy).grid(row=4,column=0,ipady=5)
    button3=tk.Button(newCanv,text='Show data',command=getstgdata).grid(row=3,column=0,ipady=5)
    
    textbox.insert(tk.END,insrtSP)
  
def save_execute_SP():
    createSP=textbox.get(1.0,tk.END)
    SP=re.search(r"(?<=PROCEDURE ).*?(?= )", createSP).group(0)
    Table=re.search(r"(?<=into ).*?(?= )", createSP).group(0)
    # Stored Procedure Drop Statement
    sqlDropSP="IF EXISTS (SELECT * FROM sys.objects \
               WHERE type='P' AND name='{}') \
               DROP PROCEDURE {}".format(SP,SP)
    
    sqlExecSP="truncate table {} \
               exec {}".format(Table,SP)

    # Drop SP if exists
    cursor.execute(sqlDropSP)
    

    # Create SP using Create statement
    cursor.execute(createSP)

    # Call SP and trap Error if raised
    try:
        cursor.execute(sqlExecSP)
        conn.commit()
        #if (stgbutton['state'] == tk.DISABLED):
            #print("Ankit")
    except Exception as e:
        print('Error !!!!! %s')% e
    
def alltables():
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG=? ",dbname)
    All_tables=sorted([item[0] for item in cursor.fetchall()],key=str.lower)   
    container = ttk.Frame(frame13,borderwidth=5,relief="solid")
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
    def chktab():
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        tablename=[(a) for a, b in checked if b == 1 ]
        for val in tablename:
            getdata(val)
            
    def droptab():
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        tablename=[(a) for a, b in checked if b == 1 ]
        for val in tablename:
            query1="DROP TABLE {}.dbo.{}".format(dbname,val)
            cursor.execute(query1)
            print(val + " Table dropped")
       
    data={}
    for x in All_tables:
        chkvar = tk.IntVar()
        C1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
        C1.pack(anchor="nw")
        DragDropListbox(C1)
        data[x] = chkvar
    
    btn = tk.Button(container, text="Data", command=chktab)
    btn.pack()
    btn1 = tk.Button(container, text="Drop", command=droptab)
    btn1.pack()
    btn2 = tk.Button(container, text="Close", command=container.destroy)
    btn2.pack()

class shiftSelectListbox(tk.Listbox):
    def __init__(self, master, **kw):
        kw['selectmode'] = tk.SINGLE
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.setCurrent)
        self.bind('<B1-Motion>', self.shiftSelection)
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

    
def allSps():
    query="SELECT ROUTINE_NAME FROM {}.INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_TYPE = 'PROCEDURE'".format(dbname)
    cursor.execute(query)
    All_Sps=sorted([item[0] for item in cursor.fetchall()],key=str.lower)   
    container = ttk.Frame(frame13,borderwidth=5,relief="solid")
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
    container.grid(row=3,column=1,padx=5,pady=5)
    frameCanvas.pack(side="left")
    scrollbar.pack(side="right", fill="y")
    
    data={}
    for x in All_Sps:
        chkvar = tk.IntVar()
        C1 = tk.Checkbutton(scrollable_frame, text=x, variable=chkvar,cursor="hand1")
        C1.pack(anchor="nw")
        data[x] = chkvar
    
    def Splist():
        global Splistbox,SpLabel,Spbutton
        SpLabel=tk.Label(frame12,text='Set SP Sequence',bg="white")
        SpLabel.grid(row=2,column=10,ipadx=20)
        Splistbox =shiftSelectListbox(frame12,height=15,  
                  width = 35,  
                  bg = "white", 
                  activestyle = 'dotbox',  
                  font = ("Helvetica",10), 
                  fg = "black")
        Splistbox.grid(row=2,column=11)
        scrollbar4 = tk.Scrollbar(frame12)
        scrollbar4.grid(row = 2,rowspan = 1, column = 12,ipady=15)
        Splistbox.config(yscrollcommand = scrollbar4.set)
        scrollbar4.config(command = Splistbox.yview)
        Spbutton=tk.Button(frame12,text='Execute',command=Spexe)
        Spbutton.grid(row=3,column=11,pady=10,ipadx=20)
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        Spname=[(a) for a, b in checked if b == 1 ]
        for val in Spname:
            Splistbox.insert(tk.END,val)
    
    def Spexe():
        Spname=Splistbox.get(0,tk.END)
        for val in Spname:
            sqlExecSP="exec {}".format(val)
            try:
                cursor.execute(sqlExecSP)
                conn.commit()
            except Exception as e:
                print('Error !!!!! %s')% e
        Splistbox.destroy()
        Spbutton.destroy()
        SpLabel.destroy()
    
    def Spedit():
        
        
        checked=[(x, chkvar.get()) for x, chkvar in data.items()]
        Spname=[(a) for a, b in checked if b == 1 ]
        for val in Spname:
            val=("'" + val + "'")
            query="SELECT ROUTINE_DEFINITION FROM {}.INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_TYPE = 'PROCEDURE' \
            and ROUTINE_NAME={}".format(dbname,val)
            cursor.execute(query)
            Spdef1=cursor.fetchall()
            for element in Spdef1:
                for item in element:
                    Spdef=item
            #Spdef=Spdef.replace("CREATE","ALTER")
            newWindow = tk.Toplevel(window)
            newCanv = tk.Canvas(newWindow, bg="white")
            newCanv.pack(expand=True)
            label3=tk.Label(newCanv,text=val,bg="white").grid(row=0,column=0)
            global textbox
            textbox=tk.Text(newCanv,bg="white",height = 20, width = 100)
            textbox.grid(row=1,column=0,padx=50,pady=50)
    
            sb = tk.Scrollbar(newCanv, orient=tk.VERTICAL, command=textbox.yview)
            sb.grid(row=1,column=1,ipady=50)
            textbox.config(yscrollcommand=sb.set)
            button1=tk.Button(newCanv,text='Save/Execute SP',command= save_execute_SP).grid(row=2,column=0,ipady=5)
            button2=tk.Button(newCanv,text='Close',command= newWindow.destroy).grid(row=4,column=0,ipady=5)
            button3=tk.Button(newCanv,text='Show data',command=getstgdata).grid(row=3,column=0,ipady=5)
            textbox.insert(tk.END,Spdef)
       
    
    
    btn = tk.Button(container, text="ListSP", command=Splist)
    btn.pack()
    btn1 = tk.Button(container, text="Edit", command=Spedit)
    btn1.pack()
    btn2 = tk.Button(container, text="Close", command=container.destroy)
    btn2.pack()



def clearall():
    
    servername=''
    dbname=''
    window.destroy()
    
button3=tk.Button(frame2,text='Show Columns',command=getlndcolumns).grid(row=2,column=5,padx=10, pady=10)
button4=tk.Button(frame9,text='Show Columns',command=getstgcolumns).pack()
button9=tk.Button(frame9,text='Map Columns',command=mapcolumns).pack()

button7=tk.Button(frame13,text='Close',command=clearall).grid(row=5,column=1,padx=10, pady=10)
button10=tk.Button(frame13,text='All Tables',command=alltables).grid(row=2,column=1,padx=10, pady=10)
button11=tk.Button(frame13,text='All SPs',command=allSps).grid(row=3,column=1,padx=10, pady=10)


#Error display
def report_callback_exception(self, exc, val, tb):
    if 'must be active, anchor, end, @x,y, or a number' in str(val):
        messagebox.showerror("Error", message='No Value Selected from List')
    if "'cursor' referenced before assignment" in str(val):
        messagebox.showerror("Error", message='Server\Database name not specified')
    else:
        messagebox.showerror("Error", message=str(val))

tk.Tk.report_callback_exception = report_callback_exception

frame1.grid(sticky=tk.E+tk.W,padx=10, pady=10)
frame2.grid(sticky=tk.E+tk.W,padx=10, pady=10)
frame3.grid(sticky=tk.E+tk.W)

frame7.grid(sticky=tk.E+tk.W)

frame8.grid()
myCanvas2.pack(expand=True,side='right')
myCanvas.pack(expand=True,side='left')

frame4.grid()




window.mainloop()

#good to have -- message box to print each activity output, separate button for save and execute
