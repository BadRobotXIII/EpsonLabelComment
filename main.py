"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
*************************************************************************************************
** Company: KayZee Automation
** Author: Kameron Zulfic
** Developer Contact: kzulfic@outlook.com
** Date:  03/03/2024
** Function: 
** Intended Use: 
*************************************************************************************************
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
#################################################################################
## Packaged with pyinstaller. Input line below to package
## command line arguments - pyinstall -noconsole -onedir EpsonLabelGeneratorCustTk.py
## Edited JSON library location C:\Users\kzulf\AppData\Local\Programs\Python\Python312\Lib\json
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
from utils import *

#Acquire python version from platform
pythonVersion = platform.python_version()

def main():
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    --|Create Window Application Driver|---------------------------------------------
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    #Declare globals
    global mainWin
    global CTbtnStart
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
    #Call window constructor class
    #initialize main window and base frame widgets
    mainWin = MainWindow()
    mainWin.app_closing = False
        
    #Set default font style and character sizing
    # fnt = tkfont.Font(family='Consolas 11', size=11, weight="normal")
    fnt = CTkFont(family='Consolas', size=16, weight='bold')
    fnt1 = CTkFont(family='Consolas', size=16, weight='normal')
    char_width = fnt.measure("0")

    """""""Configuratiion widgets"""""""
    #Checkbox - IO labels
    checkboxIOVar = IntVar(value=0)
    chkBxIOLabels = Checkbutton(mainWin.frameSetup, background='lightgrey', text='Enable IO Labels', width=16,
                                  variable=checkboxIOVar, command=check_box_toggle, font=fnt1)
    chkBxIOLabels.pack(side='top', anchor='w', padx=0, pady=0)

    #Checkbox - Position labels
    checkboxpPosVar = BooleanVar()
    chkBxPosLabels = Checkbutton(mainWin.frameSetup, background='lightgrey', text='Position Labels', width=15,
                                  variable=checkboxpPosVar, command=check_box_toggle, font=fnt1)
    chkBxPosLabels.pack(side='top', anchor='w', padx=0, pady=0)

    btnGenerate = CTkButton(mainWin.frameSetup, text='Generate Files', border_width=2, border_spacing=2, state=NORMAL,
                             fg_color='darkgrey', hover_color='grey', text_color='black', width=23, command=generate_files)
    btnGenerate.pack(anchor="w", padx=3, pady=0)

    """""""Setup  widgets"""""""    
    #Label - Save directory
    lblSaveDir = Label(mainWin.frameFileSel, justify='left', text='File Save location',
                        fg='black', bg='lightgrey',height=1, width=18, font='Consolas 11')
    lblSaveDir.pack(anchor='w', ipadx=0, ipady=0)

    #Entry - Epson project directory  
    entrySaveDirectory = StringVar()
    txbxSaveDir = CTkEntry(mainWin.frameFileSel, textvariable=entrySaveDirectory, width=450, fg_color='darkgrey',
                                    state=NORMAL, border_color='black', text_color='black', font=fnt1)
    entrySaveDirectory.set('')
    txbxSaveDir.pack(anchor='w', pady=100)

    #Button - Launch file explorer
    btnSaveDir = CTkButton(mainWin.frameFileSel, bg_color='lightgrey', border_width=2, border_color='black', fg_color='grey',
                            text_color='black', width=3, height=28, text='...', font=fnt, command=get_save_dir)
    btnSaveDir.pack(side='left', anchor=W, after=txbxSaveDir, padx=0, pady=0, ipadx=0, ipady=0)

    #Menu - Popup menu to paste from clipboard
    popupMenuTbSaveDir = Menu(txbxSaveDir, tearoff=0)
    popupMenuTbSaveDir.add_command(label='Paste', command=directory_paste)
    txbxSaveDir.bind('<Button-3>', lambda event: paste_menu(event, txbxSaveDir))
    txbxSaveDir.pack(side='left', anchor='n', padx=1, pady=1)
    
    #Label - Source directory
    lblSourceDir = Label(mainWin.frameFileSel1, justify='left', text='Source File Location',
                          fg='black', bg='lightgrey',height=1, width=20, font='Consolas 11')
    lblSourceDir.pack(anchor='w', ipadx=0, ipady=0)

    #Entry - Save directory  
    entrySourceDirectory = StringVar()
    txbxSourceDir = CTkEntry(mainWin.frameFileSel1, textvariable=entrySourceDirectory, width=450, fg_color='darkgrey',
                                    state=NORMAL,border_color='black', text_color='black', font=fnt1)
    entrySourceDirectory.set('')
    txbxSourceDir.pack(anchor='w', pady=100)
    
    #Button - Launch file explorer
    btnSourceDir = CTkButton(mainWin.frameFileSel1, bg_color='lightgrey', border_width=2, border_color='black', fg_color='grey',
                            text_color='black', width=3, height=28, text='...', font=fnt, command=get_source_dir)
    btnSourceDir.pack(side='left', anchor=W, after=txbxSourceDir, padx=0, pady=0, ipadx=0, ipady=0)

    #Menu - Popup menu to paste from clipboard
    popupMenuTbSourceDir = Menu(txbxSourceDir, tearoff=0)
    popupMenuTbSourceDir.add_command(label='Paste', command=directory_paste)
    txbxSourceDir.bind('<Button-3>', lambda event: paste_menu(event, txbxSourceDir))
    txbxSourceDir.pack(side='left', anchor='n', padx=1, pady=1)

    # add a listbox for error messages
    lbErrorMessage = CTkTextbox(mainWin.root, height=50, width=1000, fg_color='red', font=fnt1)
    lbErrorMessage.pack(anchor=S, side='left', fill=X, padx=2, pady=0)
    timeStamp = datetime.now().replace(microsecond=0)
    print(datetime.now())
    lbErrorMessage.insert(0.0, 'No errors {}'.format(datetime.now().replace(microsecond=0)))
    lbErrorMessage.insert(0.0,'\n')
    lbErrorMessage.insert(0.0, 'Error 1001 - Invalid Directory {}'.format(timeStamp))
    lbErrorMessage.insert(0.0,'\n')


    # # set the minimum window size  to the current size
    mainWin.update()

    #Call tKinter main driver
    mainWin.start()

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
---------------------|Generate files when requested by user|---------------------
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
def generate_files():
    #Declare variables and open necessary file streams
    path = os.getcwd()
    filePath = str(path)
    fileName = 'PLC Robot Interface.xlsx'
    sheetName = 'PLC > Robot'
    sheets = {'INPUT':'PLC > Robot', 'OUTPUT': 'Robot > PLC', 'ERRORS': 'Errors', 'POINTS': 'RobotPoints'}

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
        keys[r'sLabel'] = str('\"' + key + '\"')
        keys[r'sDescription'] = str('\"' + keyVal + '\"')
        print(keys)
        outFile.writelines('Label{} '.format(cell + 1))
        outFile.writelines(json.dumps(keys, sort_keys=False, indent=4, separators=('', '=')))
        outFile.writelines('\n\r')
        # Below code is correct for forming propper dictionaries to diplay propper. Commented due to Epson requireing a format different than JSON.
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
    entrySaveDirectory.set(mainWin.root.clipboard_get())
    txbxSaveDir.select_range(0, 'end')
    txbxSaveDir.icursor('end')
    return directory_paste

def paste_menu(event, tbSaveDir):
    try:
        old_clip = mainWin.clipboard_get()
    except:
        old_clip = None

    if (not old_clip is None) and (type(old_clip) is str) and tbSaveDir['state'] == 'normal':
        tbSaveDir.select_range(0, 'end')
        popupMenuTbSaveDir.post(event.x_root, event.y_root)
    
def get_save_dir():
    string = filedialog.askdirectory(parent=mainWin.root, initialdir='C:"\"',title='Please select a directory')
    entrySaveDirectory.set(string)
    print(str(string))
    return

def get_source_dir():
    string = filedialog.askdirectory(parent=mainWin.root, initialdir='C:"\"',title='Please select a directory')
    entrySourceDirectory.set(string)
    print(str(string))
    return 



if __name__=='__main__':
    main()



