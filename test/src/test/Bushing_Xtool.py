'''
Created on 20/02/2018

@author: JCHAV106
'''
from tkinter import Toplevel, Entry, Tk, ttk, Button, filedialog, messagebox, PhotoImage
from BushingsData import Extract
import sys, os

global b
global lb
b = Extract()
global text
global text2


def xml_check():
    if browser_xmlfo['state'] == 'disabled':
        browser_xmlfo.config(state='active')
        textx3.config(state='normal')
    else: 
        browser_xmlfo.config(state='disabled')
        textx3.config(state='disabled')

def exl_check():
    if browser_exlfo['state'] == 'disabled':
        browser_exlfo.config(state='active')
        texte3.config(state='normal')
    else: 
        browser_exlfo.config(state='disabled')
        texte3.config(state='disabled')
    

def run(t = None,t2 = None,t3 = None):
    if browser_exlfo['state'] == 'disabled':
        if messagebox.askokcancel(title = 'File Ouput', message = 'If you want to overwrite the original file with the updated data then press OK.\nOtherwise press CANCEL and select the XML Output File box to create a new file in the desired path.') == True:
            if t == None:
                t = texte.get()
            if t2 == None:
                t2 = texte2.get()
            if t3 == None:
                if browser_exlfo['state'] == 'disabled':
                    t3 = t2
                else:
                    t3 = texte3.get() 
        else:
            if t == None:
                t = texte.get()
            if t2 == None:
                t2 = texte2.get()
            if t3 == None:
                if browser_exlfo['state'] == 'disabled':
                    t3 = t2
                else:
                    t3 = texte3.get() 

    b.exl2xml(t, t2, t3)
    
def run2(t = None, t3 = None):
    if browser_xmlfo['state'] == 'disabled':
        if messagebox.askokcancel(title = 'File Ouput', message = 'If press OK, the Excel file with the bushings from the XML is gonna be placed in the Documents directory.\nOtherwise press CANCEL and select the Excel Output File box to create a new file in the desired path.') == True:
            if t == None:
                t = textx.get()
            if t3 == None:
                if browser_xmlfo['state'] == 'disabled':
                    t3 = None
                else:
                    t3 = textx3.get()
        else:
            if t == None:
                t = textx.get()
            if t3 == None:
                if browser_xmlfo['state'] == 'disabled':
                    t3 = None
                else:
                    t3 = textx3.get()
    b.xml2exl(t,t3)

def browserfx1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    textx.insert(0, file1)
    
def browserfx3():
    file3 = filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[("Excel file","*.xlsx")])
    textx3.insert(0,file3)

    
def browserfe1():
    file1 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    texte.insert(0, file1)
    
def browserfe2():
    file2 = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xlsx .xlsm"),("all files","*.*")))
    texte2.insert(0, file2)
    
def browserfe3():
    file3 = filedialog.asksaveasfilename(defaultextension='.xml',filetypes=[("XML file","*.xml")])
    texte3.insert(0, file3)
    
def exl2xml_win():
    newin = Toplevel(root)
    newin.configure(background = '#666666')
    display = ttk.Label(newin, text = "XML File", background = '#666666', foreground = 'white', font=('Arial',10,'bold'))
    display.grid(row=1,column=3)
    display2 = ttk.Label(newin, text = "Excel File", background = '#666666', foreground = 'white', font=('Arial',10,'bold'))
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
    texte.grid(row=1,column=4)
    texte2.grid(row=2,column=4)
    texte3.grid(row=3,column=4)
    browser_xml = Button(newin, text = "Browser", command = browserfe1, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    browser_xml.grid(row=1,column=5)
    browser_exl = Button(newin, text = "Browser", command = browserfe2, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    browser_exl.grid(row=2,column=5)
    global browser_exlfo
    browser_exlfo = Button(newin, text = "Browser", command = browserfe3, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    browser_exlfo.config(state = 'disabled', foreground = 'gray', bg = '#002F6C')
    browser_exlfo.config(state = 'normal', foreground = 'white', bg = '#002F6C')
    browser_exlfo['state'] = 'disabled'
    browser_exlfo.grid(row=3,column=5)
    style = ttk.Style()
    style.configure("Gray.TCheckbutton", background="#666666", foreground = "white", font = ('Arial',10,'bold'))
    c = ttk.Checkbutton(newin, text = "XML Output\n     File", command = exl_check, style = "Gray.TCheckbutton")
    c.grid(row=3, column=3)
    global Runne
    Runne = Button(newin, text = "Run", command = run, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    Runne.grid(row = 5, column = 5)
    return browser_exlfo, texte, texte2, texte3, Runne
    
def xml2exl_win():
    newin = Toplevel(root)
    newin.configure(background = '#666666')
    display = ttk.Label(newin, text = "XML File", background = '#666666', foreground = 'white', font=('Arial',10,'bold'))
    display.grid(row=1,column=3)
    global textx
    textx = Entry(newin)
    global textx3
    textx3 = Entry(newin,state='disabled')
    textx.grid(row=1,column=4)
    textx3.grid(row=3,column=4)
    browser_xml = Button(newin, text = "Browser", command = browserfx1, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    browser_xml.grid(row=1,column=5)
    global browser_xmlfo
    browser_xmlfo = Button(newin, text = "Browser", command = browserfx3, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    browser_xmlfo.config(state = 'disabled', foreground = 'gray', bg = '#002F6C')
    browser_xmlfo.config(state = 'normal', foreground = 'white', bg = '#002F6C')
    browser_xmlfo['state'] = 'disabled'
    browser_xmlfo.grid(row=3,column=5)
    style = ttk.Style()
    style.configure("Gray.TCheckbutton", background="#666666", foreground = 'white', font = ('Arial',10,'bold'))
    c = ttk.Checkbutton(newin, text = "   Excel\nOutput File", command = xml_check, style = "Gray.TCheckbutton" )
    c.grid(row=3, column=3)
    global Runnx
    Runnx = Button(newin, text = "Run", command = run2, foreground = 'white', bg = '#002F6C', state = 'active', activebackground = '#002F6C', activeforeground = 'white', font = ('Arial',10,'bold'))
    Runnx.grid(row = 5, column = 5)
    return browser_xmlfo, textx, textx3, Runnx
    
root = Tk()
root.title("Extract Bushing Rates Tool")
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the pyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app 
    # path into variable _MEIPASS'.
    #dirname = os.path.dirname(sys.executable)
    dirname = sys._MEIPASS
    pimage = PhotoImage(file = os.path.join(dirname,'source\Bushing_Xtool_logo.png'))
else:
    dirname = os.path.dirname(__file__)
    pimage = PhotoImage(file = os.path.join(dirname,'Bushing_Xtool_logo.png'))
            
#pimage = PhotoImage(file = os.path.join(dirname,'source\Bushing_Xtool_logo.png'))
pimage = pimage.subsample(2,2)
root.configure(background = 'white')
label = ttk.Label(root, background = 'white', image = pimage)
label.pack()
#title_label = ttk.Label(root, text='Extract Bushing Rates Tool', background = 'white', foreground = 'blue', font =("Arial", 14))
#title_label.pack()
buton = Button(root, text="Excel bushings to XML", command = exl2xml_win, fg = 'white', bg = '#002F6C', font = ('Arial',12,'bold'))
buton.place(x=200,y=280)
buton2 = Button(root, text="XML bushings to Excel", command = xml2exl_win, fg = 'white', bg = '#002F6C', font = ('Arial',12,'bold'))
buton2.place(x=200,y=330)
root.mainloop()

if __name__ == '__main__':
    pass

