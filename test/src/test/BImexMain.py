'''
Created on 20/02/2018

@author: JCHAV106
'''
import xlrd
from tkinter import Toplevel, Entry, Tk, ttk, filedialog, messagebox, Listbox
from BImex import BImex

global b
b = BImex()
global text
global text2
#global t
#global t2

def sheetse():
    x1_workbook = xlrd.open_workbook(texte2.get())
    sheet_names = x1_workbook.sheet_names()
    lb.insert(0,*sheet_names)
    
def sheetsx():
    x1_workbook = xlrd.open_workbook(textx2.get())
    sheet_names = x1_workbook.sheet_names()
    lb.insert(0,*sheet_names)

def xml_check():
    if browser_xmlfo.state() == ('disabled',):
        browser_xmlfo.config(state='enabled')
        textx3.config(state='normal')
    else: 
        browser_xmlfo.config(state='disabled')
        textx3.config(state='disabled')

def exl_check():
    if browser_exlfo.state() == ('disabled',):
        browser_exlfo.config(state='enabled')
        texte3.config(state='normal')
    else: 
        browser_exlfo.config(state='disabled')
        texte3.config(state='disabled')
    

def run(t = None,t2 = None):
    if messagebox.askokcancel(title = 'File Ouput', message = 'If you want to overwrite the original file with the updated data then press OK.\nOtherwise press CANCEL and select the XML Output File box to create a new file') == True:
        if t == None:
            t = text.get()
        if t2 == None:
            t2 = text2.get()
        b.exl2xml(t, t2, t)
    
def run2(t = None,t2 = None):
    if t == None:
        t = text.get()
    if t2 == None:
        t2 = text2.get()
    b.exl2xml(t, t2, t)

def browserfx1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    textx.insert(0, file1)
    #t = text.get()
    
def browserfx2():
    file2 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xlsx .xlsm"),("all files","*.*")))
    textx2.insert(0, file2)
    #t2 = text2.get()
    
def browserfx3():
    file3 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    textx3.insert(0, file3)
    #t2 = text2.get()
    
def browserfe1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    texte.insert(0, file1)
    #t = text.get()
    
def browserfe2():
    file2 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xlsx .xlsm"),("all files","*.*")))
    texte2.insert(0, file2)
    #t2 = text2.get()
    
def browserfe3():
    file3 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    texte3.insert(0, file3)
    #t2 = text2.get()
    
def exl2xml_win():
    newin = Toplevel(root)
    display = ttk.Label(newin, text = "XML File")
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File")
    display2.grid(row=2,column=3)
    display2 = ttk.Label(newin)
    display2.grid(row=3,column=3)
    global texte
    texte = Entry(newin)
    global texte2
    texte2 = Entry(newin)
    global texte3
    texte3 = Entry(newin)
    texte3.config(state='disabled')
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    texte.grid(row=1,column=4)
    texte2.grid(row=2,column=4)
    texte3.grid(row=3,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserfe1)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    browser_exl = ttk.Button(newin, text = "Browser", command = browserfe2)
    browser_exl.grid(row=2,column=5)
    global browser_exlfo
    browser_exlfo = ttk.Button(newin, text = "Browser", command = browserfe3, state='disabled')
    browser_exlfo.grid(row=3,column=5)
    c = ttk.Checkbutton(newin, text = "Excel Output\n     File", command = exl_check)
    c.grid(row=3, column=3)
    Runn = ttk.Button(newin, text = "Run", command = run)
    Runn.grid(row = 5, column = 5)
    sheets_b = ttk.Button(newin, text = "Get Sheet Names", command = sheetse)
    sheets_b.grid(row = 6, column = 5)
    global lb
    lb = Listbox(newin)
    lb.grid(row = 5, column = 4)
    return browser_exlfo,lb , texte, texte2, texte3
    
def xml2exl_win():
    newin = Toplevel(root)
    display = ttk.Label(newin, text = "XML File")
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File")
    display2.grid(row=2,column=3)
    global textx
    textx = Entry(newin)
    global textx2
    textx2 = Entry(newin)
    global textx3
    textx3 = Entry(newin,state='disabled')
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    textx.grid(row=1,column=4)
    textx2.grid(row=2,column=4)
    textx3.grid(row=3,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserfx1)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    browser_exl = ttk.Button(newin, text = "Browser", command = browserfx2)
    browser_exl.grid(row=2,column=5)
    global browser_xmlfo
    browser_xmlfo = ttk.Button(newin, text = "Browser", command = browserfx3, state='disabled')
    browser_xmlfo.grid(row=3,column=5)
    c = ttk.Checkbutton(newin, text = "XML Output File", command = xml_check)
    c.grid(row=3, column=3)
    Runn = ttk.Button(newin, text = "Run", command = run2)
    Runn.grid(row=5, column=5)
    sheets_b = ttk.Button(newin, text = "Get Sheet Names", command = sheetsx)
    sheets_b.grid(row = 6, column = 5)
    global lb
    lb = Listbox(newin)
    lb.grid(row = 5, column = 4)
    return browser_xmlfo,lb, textx, textx2, textx3
    
root = Tk()
root.title("BIm3x")
label = ttk.Label(root, text="BIm3x")
label.pack()
buton = ttk.Button(root, text="Excel bushings to XML", command = exl2xml_win)
buton.pack()
buton2 = ttk.Button(root, text="XML bushings to Excel", command = xml2exl_win)
buton2.pack()
#filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
#print(filename)
root.mainloop()

if __name__ == '__main__':
    pass

