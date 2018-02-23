'''
Created on 20/02/2018

@author: JCHAV106
'''
from tkinter import Toplevel, Entry, Tk, ttk, filedialog
from BImex import BImex
from IPython.lib.clipboard import tkinter_clipboard_get

global b
b = BImex()
global text
global text2
global text3
#global t
#global t2

def run(t = None,t2 = None):
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

def browserf1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    text.insert(0, file1)
    #t = text.get()
    
def browserf2():
    file2 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    text2.insert(0, file2)
    #t2 = text2.get()
    
def browserf3():
    file3 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    text3.insert(0, file3)
    #t2 = text2.get()
    
def exl2xml_win():
    newin = Toplevel(root)
    display = ttk.Label(newin, text = "XML File")
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File")
    display2.grid(row=2,column=3)
    display2 = ttk.Label(newin)
    display2.grid(row=3,column=3)
    text = Entry(newin)
    text2 = Entry(newin)
    text3 = Entry(newin)
    text3.config(state='disabled')
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    text.grid(row=1,column=4)
    text2.grid(row=2,column=4)
    text3.grid(row=3,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserf1)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    browser_exl = ttk.Button(newin, text = "Browser", command = browserf2)
    browser_exl.grid(row=2,column=5)
    browser_exlfo = ttk.Button(newin, text = "Browser", command = browserf3, state='disabled')
    browser_exlfo.grid(row=3,column=5)
    Met = ttk.Button(newin, text = "Run", command = run)
    Met.grid(row = 4, column = 5)
    
def xml2exl_win():
    newin = Toplevel(root)
    display = ttk.Label(newin, text = "XML File")
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File")
    display2.grid(row=2,column=3)
    text = Entry(newin)
    text2 = Entry(newin)
    text3 = Entry(newin,state='disabled')
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    text.grid(row=1,column=4)
    text2.grid(row=2,column=4)
    text3.grid(row=3,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserf1)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    browser_exl = ttk.Button(newin, text = "Browser", command = browserf2)
    browser_exl.grid(row=2,column=5)
    browser_xmlfo = ttk.Button(newin, text = "Browser", command = browserf3, state='disabled')
    browser_xmlfo.grid(row=3,column=5)
    c = ttk.Checkbutton(newin, text = "XML Output File" )
    c.grid(row=3, column=3)
    Runn = ttk.Button(newin, text = "Run", command = run2)
    Runn.grid(row=5, column=5)
    
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

