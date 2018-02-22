'''
Created on 20/02/2018

@author: JCHAV106
'''
from tkinter import Toplevel
from tkinter import Entry
from tkinter import Tk
from tkinter import ttk
from tkinter import filedialog
from BImex import BImex

global b
b = BImex()
global t
global t2

def a():
    b.exl2xml('C:\\Users\\JCHAV106\\git\\Scripts\\test\\src\\test_data\\output2.xml', 'C:\\Users\\JCHAV106\\git\\Scripts\\test\\src\\test_data\\Extract_bushing_rates_tool.xlsm', 'C:\\Users\\JCHAV106\\git\\Scripts\\test\\src\\test_data\\output4.xml')

def browserf1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    text.insert(0, file1)
    t = text.get()
    
def browserf2():
    file2 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    text2.insert(0, file2)
    t2 = text2.get()
    
def exl2xml_win():
    newin = Toplevel(root)
    display = ttk.Label(newin, text = "XML File")
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File")
    display2.grid(row=2,column=3)
    global text
    text = Entry(newin)
    global text2 
    text2 = Entry(newin)
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    text2.grid(row=2,column=4)
    text.grid(row=1,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserf1)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    browser_exl = ttk.Button(newin, text = "Browser", command = browserf2)
    browser_exl.grid(row=2,column=5)
    Met = ttk.Button(newin, text = "Run", command = a)
    Met.grid(row = 3, column = 6)


root = Tk()
root.title("BIm3x")
label = ttk.Label(root, text="BIm3x")
label.pack()
buton = ttk.Button(root, text="Excel bushings to XML", command = exl2xml_win)
buton.pack()
buton2 = ttk.Button(root, text="XML bushings to Excel")
buton2.pack()
#filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
#print(filename)
root.mainloop()

if __name__ == '__main__':
    pass

