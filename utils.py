import platform
from tkinter import *
from tkinter import filedialog  
import tkinter.font as tkfont
from customtkinter import *

class MainWindow:
    def __init__(self):
        self.root = CTk()
        # windowGeo = str(winWidth) + 'x' + str(winHeight)
        try:
            self.root.iconbitmap('atomic-16.ico')
        except:
            print('Using default image')

        self.root.config(background='#222222')
        self.root._set_appearance_mode('dark')
        self.root.title('Epson Label Generator v1.0.0 ' + 'Python v ' + platform.python_version())
        self.root.geometry('500x500')
        self.root.bind('<Destroy>', self.on_exit)

                #Frame to hold configuration variables
        self.frameSetup = Frame(self.root, background='Lightgrey', width=500, height=200)
        self.frameSetup.pack(anchor='w', fill=X, ipadx=0, ipady=0)
        #Frame to hold file selection properties
        self.frameFileSel = Frame(self.root, background='Lightgrey', width=500, height=200)
        self.frameFileSel.pack(anchor='e', fill=X, ipadx=0, ipady=0)

        #Frame to hold file selection properties
        self.frameFileSel1 = Frame(self.root, background='Lightgrey', width=500, height=200)
        self.frameFileSel1.pack(anchor='e', fill=X, ipadx=0, ipady=0)

    def start(self):
        self.root.mainloop()

    def update(self):
        # set the minimum window size  to the current size
        self.root.update()
        self.root.minsize(self.root.winfo_width(), self.root.winfo_height() * .75)

    def on_exit(self):

        self.app_closing = True
    
# width wise resizing of the tag label (window)
class LabelResizing(Label):
    def __init__(self,parent,**kwargs):
        Label.__init__(self,parent,**kwargs)
        self.bind("<Configure>", self.on_resize)
        self.width = self.winfo_reqwidth()

    def on_resize(self,event):
        if self.width > 0:
            self.width = int(event.width)
            self.config(width=self.width, wraplength=self.width)

# width wise resizing of the tag entry box (window)
class EntryResizing(Entry):
    def __init__(self,parent,**kwargs):
        Entry.__init__(self,parent,**kwargs)
        self.bind("<Configure>", self.on_resize)
        self.width = self.winfo_reqwidth()

    def on_resize(self,event):
        if self.width > 0:
            self.width = int(event.width)
            self.config(width=self.width)