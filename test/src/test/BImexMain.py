'''
Created on 20/02/2018

@author: JCHAV106
'''
from tkinter import Toplevel, Entry, Tk, ttk, filedialog, messagebox, Listbox, PhotoImage
from BushingsData import Extract

global b
global lb
b = Extract()
global text
global text2
#global t
#global t2

'''def sheetse():
    x1_workbook = xlrd.open_workbook(texte2.get())
    sheet_names = x1_workbook.sheet_names()
    lb.insert(0,*sheet_names)
    

def sheetsx():
        Runnx.config(state='enabled')
        x1_workbook = xlrd.open_workbook(textx2.get())
        sheet_names = x1_workbook.sheet_names()
        lb.insert(0,*sheet_names)'''


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
    

def run(t = None,t2 = None,t3 = None):
    if messagebox.askokcancel(title = 'File Ouput', message = 'If you want to overwrite the original file with the updated data then press OK.\nOtherwise press CANCEL and select the XML Output File box to create a new file') == True:
        if t == None:
            t = texte.get()
        if t2 == None:
            t2 = texte2.get()
        if t3 == None:
            if browser_exlfo.state() == ('disabled',):
                t3 = t2
            else:
                t3 = texte3.get() 

    b.exl2xml(t, t2, t3)
    
def run2(t = None, t3 = None):
    if t == None:
        t = textx.get()
    if t3 == None:
        if browser_xmlfo.state() == ('disabled',):
            t3 = None
        else:
            t3 = textx3.get()
    b.xml2exl(t,t3)

def browserfx1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    textx.insert(0, file1)
    #t = text.get()
    
def browserfx3():
    #global file3x
    file3 = filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[("Excel file","*.xlsx")])
    textx3.insert(0,file3)
    #return file3x
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
    file3 = filedialog.asksaveasfilename(defaultextension='.xml',filetypes=[("XML file","*.xml")])
    texte3.insert(0, file3)
    #t2 = text2.get()
    
def exl2xml_win():
    newin = Toplevel(root)
    newin.configure(background = 'white')
    display = ttk.Label(newin, text = "XML File", background = 'white')
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File", background = 'white')
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
    style = ttk.Style()
    style.configure("White.TCheckbutton", background="white")
    c = ttk.Checkbutton(newin, text = "XML Output\n     File", command = exl_check, style = "White.TCheckbutton")
    c.grid(row=3, column=3)
    global Runne
    Runne = ttk.Button(newin, text = "Run", command = run,state='enabled')
    Runne.grid(row = 5, column = 5)
    return browser_exlfo, texte, texte2, texte3, Runne
    
def xml2exl_win():
    newin = Toplevel(root)
    newin.configure(background = 'white')
    display = ttk.Label(newin, text = "XML File", background = 'white')
    display.grid(row=1,column=3)
    global textx
    textx = Entry(newin)
    global textx3
    textx3 = Entry(newin,state='disabled')
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    textx.grid(row=1,column=4)
    #textx2.grid(row=2,column=4)
    textx3.grid(row=3,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserfx1)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    #browser_exl = ttk.Button(newin, text = "Browser", command = browserfx2)
    #browser_exl.grid(row=2,column=5)
    global browser_xmlfo
    browser_xmlfo = ttk.Button(newin, text = "Browser", command = browserfx3, state='disabled')
    browser_xmlfo.grid(row=3,column=5)
    style = ttk.Style()
    style.configure("White.TCheckbutton", background="white")
    c = ttk.Checkbutton(newin, text = "   Excel\nOutput File", command = xml_check, style = "White.TCheckbutton" )
    c.grid(row=3, column=3)
    global Runnx
    Runnx = ttk.Button(newin, text = "Run", command = run2, state='enabled')
    Runnx.grid(row = 5, column = 5)
    return browser_xmlfo, textx, textx3, Runnx
    
root = Tk()
root.title("Extract Bushing Rates Tool")
pimage = PhotoImage(file = 'C:\\Users\\JCHAV106\\Pictures\\ford-logo-A8C4E442AE-seeklogo.com.png')
root.configure(background = 'white')
label = ttk.Label(root, background = 'white', image = pimage)
label.pack()
title_label = ttk.Label(root, text='Extract Bushing Rates Tool', background = 'white', foreground = 'blue', font =("Arial", 14))
title_label.pack()
buton = ttk.Button(root, text="Excel bushings to XML", command = exl2xml_win)
buton.pack()
buton2 = ttk.Button(root, text="XML bushings to Excel", command = xml2exl_win)
buton2.pack()
#filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
#print(filename)
root.mainloop()

if __name__ == '__main__':
    pass

