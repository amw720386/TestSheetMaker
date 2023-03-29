import tkinter as tk
import tkinter.filedialog as fd
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageWin
import win32ui
import win32con
import win32print
from pathlib import Path
import os

path = str(Path(__file__).parent.absolute())
root = tk.Tk()
root.title('TestSheetMaker v0.0.1')
root.state('zoomed')
sheets = ['']
projectFont = 'Candara'


def errorDisplay(textDisplay):
    Error = tk.Toplevel(root)
    Error.title('ERROR')
    Error.geometry('200x75')
    Error.focus_set()
    Error.grab_set()
    ErrorText = tk.Label(Error,
                         text=textDisplay,
                         font=projectFont)
    ErrorText.place(relx=0.5, rely=0.5, anchor='center')


def sheetImgUpdate():
    sheetImg = Image.open(currentSheetDir)
    sheetImgDraw = ImageDraw.Draw(sheetImg)
    sheetImgFont = ImageFont.truetype(projectFont, 24)

    for i, tb in enumerate(renderedTbs['Participants']):
        if not str(tb.get()):
            continue
        else:
            sheetImgDraw.text((trueCurrentUnscaledTbPos['Participants'][i][0], trueCurrentUnscaledTbPos['Participants'][i][1]), str(
                tb.get()), font=sheetImgFont, fill='black')

    for i in ['Instructor', 'Session', 'Date', 'Location', 'CourseCode']:
        if not str(renderedTbs[i].get()):
            continue
        else:
            sheetImgDraw.text((trueCurrentUnscaledTbPos[i][0], trueCurrentUnscaledTbPos[i][1]), str(
                renderedTbs[i].get()), font=sheetImgFont, fill='black')

    def getSaved():
        sheetNameText = sheetName.get()

        if sheetNameText.strip() == '':
            errorDisplay('ENTER FILE NAME')
            return

        savePath = fd.askdirectory()

        if os.path.isfile(savePath + f'/{sheetNameText}.pdf'):
            ind = 1
            while True:
                if not os.path.isfile(savePath + f'/{sheetNameText}({ind}).pdf'):
                    sheetImg.save(savePath + f'/{sheetNameText}({ind}).pdf')
                    break
                ind += 1
        else:
            sheetImg.save(savePath + f'/{sheetNameText}.pdf')

        saveScreen.destroy()
        global sheetNamePath
        sheetNamePath = savePath + f'/{sheetNameText}.pdf'

    def getPrinted():
        getSaved()

        printers = win32print.EnumPrinters(1)

        hdc = win32ui.CreateDC

        printScreen = tk.Toplevel(root)
        printScreen.title('Print')
        printScreen.geometry('400x150')
        printScreen.focus_set()
        printScreen.grab_set()

    saveScreen = tk.Toplevel(root)
    saveScreen.title('Export')
    saveScreen.geometry('400x150')
    saveScreen.focus_set()
    saveScreen.grab_set()
    sheetNameText = tk.Label(saveScreen, text='Enter Name', font=projectFont)
    sheetNameText.place(relx=0.5, rely=0.5, anchor='center')
    sheetName = tk.Entry(saveScreen)
    sheetName.place(relx=0.5, rely=0, anchor='n', width=300, height=30)
    sheetSaveButton = tk.Button(saveScreen,
                                font=projectFont,
                                activebackground='lightblue',
                                activeforeground='white',
                                text="Save",
                                bd="2",
                                relief='raised',
                                padx=5, pady=3,
                                command=getSaved)
    sheetSaveButton.place(relx=0.5, rely=1, anchor='se')
    sheetPrintButton = tk.Button(saveScreen,
                                 font=projectFont,
                                 activebackground='lightblue',
                                 activeforeground='white',
                                 text="Print",
                                 bd="2",
                                 relief='raised',
                                 padx=5, pady=3,
                                 command=getPrinted)
    sheetPrintButton.place(relx=0.5, rely=1, anchor='sw')


def tbRender(type, level):
    global trueCurrentUnscaledTbPos, renderedTbs
    if not type == 'Preschool':
        level = int(level)
    noOfParticipants = 0
    if type == 'SFL':
        if level < 6:
            noOfParticipants = 6
        elif level == 6:
            noOfParticipants = 8
        elif level < 9:
            noOfParticipants = 10
        else:
            noOfParticipants = 12
    noOfParticipants = 12 if type == 'Parent & Tot' else noOfParticipants
    noOfParticipants = 4 if level in [
        'A', 'B'] and type == 'Preschool' else 5 if type == 'Preschool' else noOfParticipants
    noOfParticipants = 12 if type == 'Adult' else noOfParticipants
    if noOfParticipants == 0:
        errorDisplay('How did we get here?')

    unscaledTbPos = {4: {'Participants': [[140, 639], [140, 872], [140, 1106], [140, 1339]], 'Instructor': [205, 365], 'Session': [250, 425], 'Date': [220, 485], 'Location': [180, 545], 'CourseCode': [1964, 60]},
                     5: {'Participants': [[140, 655], [140, 840], [140, 1024], [140, 1208], [140, 1392]], 'Instructor': [205, 365], 'Session': [250, 425], 'Date': [220, 485], 'Location': [180, 545], 'CourseCode': [1964, 60]},
                     6: {'Participants': [[130, 656], [130, 803], [130, 950], [130, 1098], [130, 1245], [130, 1392]], 'Instructor': [205, 365], 'Session': [250, 425], 'Date': [220, 485], 'Location': [180, 545], 'CourseCode': [1964, 60]},
                     8: {'Participants': [[130, 637], [130, 752], [130, 868], [130, 984], [130, 1099], [130, 1215], [130, 1330], [130, 1446]], 'Instructor': [205, 365], 'Session': [250, 425], 'Date': [220, 485], 'Location': [180, 545], 'CourseCode': [1964, 60]},
                     10: {'Participants': [[140, 619], [140, 712], [140, 804], [140, 896], [140, 989], [140, 1081], [140, 1174], [140, 1266], [140, 1358], [140, 1451]], 'Instructor': [205, 365], 'Session': [250, 425], 'Date': [220, 485], 'Location': [180, 545], 'CourseCode': [1964, 60]},
                     12: {'Participants': [[140, 600], [140, 677], [140, 754], [140, 832], [140, 909], [140, 986], [140, 1063], [140, 1140], [140, 1217], [140, 1294], [140, 1371], [140, 1448]], 'Instructor': [205, 365], 'Session': [250, 425], 'Date': [220, 485], 'Location': [180, 545], 'CourseCode': [1964, 60]}}

    currentUnscaledTbPos = unscaledTbPos[noOfParticipants]
    trueCurrentUnscaledTbPos = currentUnscaledTbPos.copy()
    trueCurrentUnscaledTbPos['Participants'] = []
    for i in currentUnscaledTbPos['Participants']:
        trueCurrentUnscaledTbPos['Participants'].append(i.copy())

    scaledTbPos = {'Participants': [], 'Instructor': [], 'Session': [
    ], 'Date': [], 'Location': [], 'CourseCode': []}
    renderedTbs = {'Participants': [], 'Instructor': tk.Entry(root), 'Session': tk.Entry(
        root), 'Date': tk.Entry(root), 'Location': tk.Entry(root), 'CourseCode': tk.Entry(root)}

    for coord in currentUnscaledTbPos['Participants']:
        coord[0] = int((sheetDisplayRes[0]/2200)*coord[0])
        coord[1] = int((sheetDisplayRes[1]/1700)*coord[1])
        scaledTbPos['Participants'].append(coord)

    for i in ['Instructor', 'Session', 'Date', 'Location', 'CourseCode']:
        scaledTbPos[i] = [int((sheetDisplayRes[0]/2200)*currentUnscaledTbPos[i][0]),
                          int((sheetDisplayRes[1]/1700)*currentUnscaledTbPos[i][1])]

    for _ in scaledTbPos['Participants']:
        renderedTbs['Participants'].append(tk.Entry(root))

    for tb, tbPos in zip(renderedTbs['Participants'], scaledTbPos['Participants']):
        DisplayCanvas.create_window(tbPos[0], tbPos[1], anchor='nw', width=(
            (sheetDisplayRes[0]/2200)*215), height=((sheetDisplayRes[1]/1700)*32), window=tb)

    for i in ['Instructor', 'Session', 'Date', 'Location', 'CourseCode']:
        DisplayCanvas.create_window(scaledTbPos[i][0], scaledTbPos[i][1], anchor='nw', width=(
            (sheetDisplayRes[0]/2200)*215), height=((sheetDisplayRes[1]/1700)*32), window=renderedTbs[i])


def displaySheet():
    global sheetDisplay, sheetDisplayRes, DisplayCanvas, currentSheetDir
    if newfileButton.winfo_exists:
        newfileButton.destroy()
    E1 = tk.Entry(root, bd=5)
    currentSheet = sheets[0]
    currentSheetDir = f'{path}\\images\\'
    currentSheetDir += 'PARENTANDTOT\\' if currentSheet[0] == 'Parent & Tot' else currentSheet[0].upper(
    ) + '\\'
    currentSheetDir += 'ParentTot' if currentSheet[0] == 'Parent & Tot' else currentSheet[0]
    currentSheetDir += str(currentSheet[1]) + '.jpg'
    sheetDisplayI = Image.open(currentSheetDir)
    sheetDisplayRes = [sheetDisplayI.width, sheetDisplayI.height]

    if abs(root.winfo_width() - sheetDisplayI.width) > abs(root.winfo_height() - sheetDisplayI.height):
        sheetDisplayRes[0] = root.winfo_width() - 100
        sheetDisplayRes[1] = int(sheetDisplayRes[0]/1.29411765)
    else:
        sheetDisplayRes[1] = root.winfo_height() - 100
        sheetDisplayRes[0] = int(sheetDisplayRes[1] * 1.29411765)

    sheetDisplayI = sheetDisplayI.resize((
        sheetDisplayRes[0], sheetDisplayRes[1]))
    sheetDisplay = ImageTk.PhotoImage(sheetDisplayI)
    DisplayCanvas = tk.Canvas(
        root, width=sheetDisplayRes[0], height=sheetDisplayRes[1])
    DisplayCanvas.place(relx=0.5, rely=0.5, anchor='center')
    DisplayCanvas.create_image(0,
                               0, anchor='nw', image=sheetDisplay)

    tbRender(currentSheet[0], currentSheet[1])

    changeSheetButton = tk.Button(root,
                                  font=projectFont,
                                  activebackground='lightblue',
                                  activeforeground='white',
                                  text="Change Sheet",
                                  bd="2",
                                  relief='raised',
                                  padx=15, pady=12,
                                  command=createNewSheet)
    changeSheetButton.place(relx=1, rely=0, anchor='ne')

    saveButton = tk.Button(root,
                           font=projectFont,
                           activebackground='lightblue',
                           activeforeground='white',
                           text="Export",
                           bd="2",
                           relief='raised',
                           padx=15, pady=12,
                           command=sheetImgUpdate)
    saveButton.place(relx=1, rely=1, anchor='se')


def createNewSheet():
    newSheetScreen = tk.Toplevel(root)
    newSheetScreen.title('New Sheet')
    newSheetScreen.geometry('500x500')
    newSheetScreen.focus_set()
    newSheetScreen.grab_set()

    def setSheet():
        if str(setLevel.get()) == '':
            errorDisplay('LEVEL NOT SELECTED')
        else:
            sheets[0] = ([str(setClass.get()), str(setLevel.get())])
            newSheetScreen.destroy()
            displaySheet()

    def updateLevels(*args):
        levelOptionList["menu"].delete(0, "end")
        setLevel.set('')
        for string in typesOfLevels[str(setClass.get())]:
            levelOptionList['menu'].add_command(
                label=string, command=tk._setit(setLevel, string))

    typesOfClasses = [
        'SFL',
        'Adult',
        'Parent & Tot',
        'Preschool'
    ]
    setClass = tk.StringVar()
    setClass.set('SFL')
    classOptionList = tk.OptionMenu(newSheetScreen, setClass, *typesOfClasses)
    classOptionList.place(relx=0.5, rely=0.5, anchor='center')
    setClassLabel = tk.Label(newSheetScreen,
                             text='What type of class do you want?',
                             font=projectFont)
    setClassLabel.place(relx=0.5, rely=0.4, anchor='n')

    typesOfLevels = {'SFL': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11],
                     'Adult': [1, 2, 3, 4, 5],
                     'Parent & Tot': [1, 2, 3],
                     'Preschool': ['A', 'B', 'C', 'D', 'E', 'F']}
    setLevel = tk.StringVar()
    setLevel.set('')
    levelOptionList = tk.OptionMenu(
        newSheetScreen, setLevel, *typesOfLevels[str(setClass.get())])
    levelOptionList.place(relx=0.5, rely=0.7, anchor='center')
    setLevelLabel = tk.Label(newSheetScreen,
                             text='What level do you want?',
                             font=projectFont)
    setLevelLabel.place(relx=0.5, rely=0.6, anchor='n')

    setClass.trace('w', updateLevels)

    donefileButton = tk.Button(newSheetScreen,
                               font=projectFont,
                               activebackground='lightblue',
                               activeforeground='white',
                               text="Done",
                               bd="2",
                               relief='raised',
                               padx=15, pady=12,
                               command=setSheet)
    donefileButton.place(relx=0.5, rely=1.0, anchor='s')


root.configure(bg='grey')
newfileButton = tk.Button(root,
                          font=projectFont,
                          activebackground='lightblue',
                          activeforeground='white',
                          text="New Sheet",
                          bd="2",
                          relief='raised',
                          padx=15, pady=12,
                          command=createNewSheet)
newfileButton.place(relx=0.5, rely=0.5, anchor='center')


root.mainloop()
