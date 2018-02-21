'''
Created on 20/02/2018

@author: JCHAV106
'''
from tkinter import * 
from tkinter import ttk
from tkinter import filedialog

def browserf():
    filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))

def exl2xml_win():
    newin = Toplevel(root)
    display = ttk.Label(newin, text = "XML File")
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File")
    display2.grid(row=2,column=3)
    text = Entry(newin)
    #text.pack(side=LEFT, fill = X)
    #Grid.columnconfigure(newin,4,weight=2)
    text.grid(row=1,column=4)
    browser_xml = ttk.Button(newin, text = "Browser", command = browserf)
    browser_xml.grid(row=1,column=5)
    #browser.pack(side=RIGHT)
    browser_exl = ttk.Button(newin, text = "Browser", command = browserf)
    browser_exl.grid(row=2,column=5)


root = Tk()
root.title("BIm3x")
label = ttk.Label(root, text="BIm3x")
label.pack()
buton = ttk.Button(root, text="Excel bushings to XML", command = exl2xml_win)
buton.pack()
buton = ttk.Button(root, text="XML bushings to Excel")
buton.pack()
#filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
#print(filename)
root.mainloop()

if __name__ == '__main__':
    pass

