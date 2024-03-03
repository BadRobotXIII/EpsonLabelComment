"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
 __    __    _______
|%%\  / %%\ /%%%%%%%\ 
|%% |/ %% /\____%% / 
|%% |%%  /     %%  / 
|%%%%%  /     %%  /  
|%%  %%\     %%  /   
|%% |\%%\   %%  /    
|%% | \%%\ %%%%%%%%\ 
 \__|  \__|\________|
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
#################################################################################
## Packaged with pyinstaller. Input line below to package
## command line arguments - pyinstall -noconsole -onedir EpsonLabelGeneratorCustTk.py
#################################################################################

import os
import openpyxl
import time
import json
import platform
from datetime import datetime
from tkinter import *
from tkinter import filedialog  
import tkinter.font as tkfont
from customtkinter import *
import pylogix

pythonVersion = platform.python_version()

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
def main():
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    --|Create Window Application Driver|---------------------------------------------
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    #Declare globals
    global root
    global CTbtnStart
    global app_closing
    global entrySaveDirectory
    global entrySourceDirectory
    global txbxSaveDir
    global txbxSourceDir
    global popupMenuTbSaveDir
    global popupMenuTbSourceDir
    global lbErrorMessage

    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    --|tKinter User Interface|-------------------------------------------------------
    --|Create  Window Application Driver|--------------------------------------------
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    #Instantiate base window
    root = CTk()
    try:
        root.iconbitmap('atomic-16.ico')
    except:
        print('Using default image')
    root.config(background='#222222')
    root._set_appearance_mode('dark')
    root.title('Epson Label Generator v1.0.0 ' + 'Python v ' + pythonVersion)
    root.geometry('500x500')
    root.bind('<Destroy>', on_exit)
    app_closing = False

    #Set default font style and character sizing
    # fnt = tkfont.Font(family='Consolas 11', size=11, weight="normal")
    fnt = CTkFont(family='Consolas', size=16, weight='bold')
    fnt1 = CTkFont(family='Consolas', size=16, weight='normal')
    char_width = fnt.measure("0")

    """""""Initialize frames"""""""
        #Frame to hold configuration variables
    frameSetup = Frame(root, background='Lightgrey', width=500, height=200)
    frameSetup.pack(anchor='w', fill=X, ipadx=0, ipady=0)
    #Frame to hold file selection properties
    frameFileSel = Frame(root, background='Lightgrey', width=500, height=200)
    frameFileSel.pack(anchor='e', fill=X, ipadx=0, ipady=0)

    #Frame to hold file selection properties
    frameFileSel1 = Frame(root, background='Lightgrey', width=500, height=200)
    frameFileSel1.pack(anchor='e', fill=X, ipadx=0, ipady=0)

    """""""Configuratiion widgets"""""""
    #Checkbox - IO labels
    checkboxIOVar = IntVar(value=0)
    chkBxIOLabels = Checkbutton(frameSetup, background='lightgrey', text='Enable IO Labels', width=16,
                                  variable=checkboxIOVar, command=check_box_toggle, font=fnt1)
    chkBxIOLabels.pack(side='top', anchor='w', padx=0, pady=0)

    #Checkbox - Position labels
    checkboxpPosVar = BooleanVar()
    chkBxPosLabels = Checkbutton(frameSetup, background='lightgrey', text='Position Labels', width=15,
                                  variable=checkboxpPosVar, command=check_box_toggle, font=fnt1)
    chkBxPosLabels.pack(side='top', anchor='w', padx=0, pady=0)
    CTkFont

    btnGenerate = CTkButton(frameSetup, text='Generate Files', border_width=2, border_spacing=2, state=NORMAL,
                             fg_color='darkgrey', hover_color='grey', text_color='black', width=23, command=generate_files)
    btnGenerate.pack(anchor="w", padx=3, pady=0)

    """""""Setup  widgets"""""""    
    #Label - Save directory
    lblSaveDir = Label(frameFileSel, justify='left', text='File Save location',
                        fg='black', bg='lightgrey',height=1, width=18, font='Consolas 11')
    lblSaveDir.pack(anchor='w', ipadx=0, ipady=0)

    #Entry - Save directory  
    entrySaveDirectory = StringVar()
    txbxSaveDir = CTkEntry(frameFileSel, textvariable=entrySaveDirectory, width=450, fg_color='darkgrey',
                                    state=NORMAL, border_color='black', text_color='black', font=fnt1)
    entrySaveDirectory.set(r'C:\EpsonRC70\projects')
    txbxSaveDir.pack(anchor='w', pady=100)

    #Button - Launch file explorer
    btnSaveDir = CTkButton(frameFileSel, bg_color='lightgrey', border_width=2, border_color='black', fg_color='grey',
                            text_color='black', width=3, height=28, text='...', font=fnt, command=get_directory_loc)
    btnSaveDir.pack(side='left', anchor=W, after=txbxSaveDir, padx=0, pady=0, ipadx=0, ipady=0)

    #Menu - Popup menu to paste from clipboard
    popupMenuTbSaveDir = Menu(txbxSaveDir, tearoff=0)
    popupMenuTbSaveDir.add_command(label='Paste', command=directory_paste)
    txbxSaveDir.bind('<Button-3>', lambda event: paste_menu(event, txbxSaveDir))
    txbxSaveDir.pack(side='left', anchor='n', padx=1, pady=1)
    
    #Label - Source directory
    lblSourceDir = Label(frameFileSel1, justify='left', text='Source File Location',
                          fg='black', bg='lightgrey',height=1, width=20, font='Consolas 11')
    lblSourceDir.pack(anchor='w', ipadx=0, ipady=0)

    #Entry - Save directory  
    entrySourceDirectory = StringVar()
    txbxSourceDir = CTkEntry(frameFileSel1, textvariable=entrySourceDirectory, width=450, fg_color='darkgrey',
                                    state=NORMAL,border_color='black', text_color='black', font=fnt1)
    entrySourceDirectory.set(r'C:\EpsonRC70\projects')
    txbxSourceDir.pack(anchor='w', pady=100)
    
    #Button - Launch file explorer
    btnSourceDir = CTkButton(frameFileSel1, bg_color='lightgrey', border_width=2, border_color='black', fg_color='grey',
                            text_color='black', width=3, height=28, text='...', font=fnt, command=get_directory_loc)
    btnSourceDir.pack(side='left', anchor=W, after=txbxSourceDir, padx=0, pady=0, ipadx=0, ipady=0)

    #Menu - Popup menu to paste from clipboard
    popupMenuTbSourceDir = Menu(txbxSourceDir, tearoff=0)
    popupMenuTbSourceDir.add_command(label='Paste', command=directory_paste)
    txbxSourceDir.bind('<Button-3>', lambda event: paste_menu(event, txbxSourceDir))
    txbxSourceDir.pack(side='left', anchor='n', padx=1, pady=1)

    # add a listbox for error messages
    lbErrorMessage = CTkTextbox(root, height=50, width=1000, fg_color='red', font=fnt1)
    lbErrorMessage.pack(anchor=S, side='left', fill=X, padx=2, pady=0)
    timeStamp = datetime.now().replace(microsecond=0)
    print(datetime.now())
    lbErrorMessage.insert(0.0, 'No errors {}'.format(datetime.now().replace(microsecond=0)))
    lbErrorMessage.insert(0.0,'\n')
    lbErrorMessage.insert(0.0, 'Error 1001 - Invalid Directory {}'.format(timeStamp))
    lbErrorMessage.insert(0.0,'\n')


    # set the minimum window size  to the current size
    root.update()
    root.minsize(root.winfo_width(), root.winfo_height() * .75) 

    #Call tKinter main driver
    root.mainloop()

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
---------------------|Generate files when requested by user|---------------------
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
def generate_files():
    #Declare variables and open necessary file streams
    filePath = ''
    filePath = r'C:\Users\kzulf\Dropbox\Coding\Python\EpsonComment'
    fileName = 'PLC Robot Interface.xlsx'
    sheetName = 'PLC > Robot'

    filePath  = os.path.join(filePath, fileName)
    print(filePath)

    startTime = time.time()
    print('Start time {}'.format(startTime))
    lbErrorMessage.insert(0.0, 'Start time {}\n'.format(startTime))

    #Open workbook and set active sheet
    inFile = openpyxl.load_workbook(filePath)
    if os.path.exists('IOLabels.dat'):
        os.remove('IOLabels.dat')

    outFile = open('IOLabels.dat', 'x')

    inFile.active = inFile[sheetName]
    sheet = inFile.active
    keys = {}
    for cell in range(len(sheet['A'])):
        key =  str(sheet['A'][cell].value).format(cell + 1)
        keyVal = str(sheet['B'][cell].value)
        keys[r'nbit'] = str(cell)
        keys[r'sLabel'] = str(key)
        keys[r'sDescription'] = str(keyVal)
        outFile.writelines('Label{} '.format(cell + 1))
        outFile.writelines(json.dumps(keys, sort_keys=False, indent=4, separators=('', '=')))
        outFile.writelines('\n\r')
        #Below code is the correct for forming propper dictionaries to diplay propper. Commented due to Epson requireing a format different than JSON.
        # labelObj = {'Label{} '.format(cell+1): labelDict}
        # jsonObj = json.dumps(labelDict, sort_keys=False, indent=4, separators=('', '='))
        # outFile.writelines(json.dumps(labelObj, sort_keys=False, indent=4, separators=('', '=')))
    endTime = time.time()
    executionTIme = (endTime - startTime)
    print('End time {}'.format(endTime))
    lbErrorMessage.insert(0.0, 'End time {}\n'.format(endTime))
    print('Execution time {}'.format(executionTIme))
    lbErrorMessage.insert(0.0, 'Execution time {}\n'.format(executionTIme))

#Toggle variable on check and uncheck of check box object
def check_box_toggle(*args):
    print('Checkbox Checked')
    return

def directory_paste():
    # user clicked the 'Paste' option so paste the IP Address from the clipboard
    entrySaveDirectory.set(root.clipboard_get())
    txbxSaveDir.select_range(0, 'end')
    txbxSaveDir.icursor('end')
    return directory_paste

def paste_menu(event, tbSaveDir):
    try:
        old_clip = root.clipboard_get()
    except:
        old_clip = None

    if (not old_clip is None) and (type(old_clip) is str) and tbSaveDir['state'] == 'normal':
        tbSaveDir.select_range(0, 'end')
        popupMenuTbSaveDir.post(event.x_root, event.y_root)
    
def get_directory_loc():
    string = filedialog.askdirectory(parent=root, initialdir='C:"\"',title='Please select a directory')
    entrySaveDirectory.set(string)
    print(str(string))
    return 

def on_exit(*args):
    global app_closing

    app_closing = True

if __name__=='__main__':
    main()



