from pyautogui import *
import pyautogui
import time
import keyboard
import pandas as pd
import xlsxwriter
import openpyxl
import sys, os #till here is SHREYAS imports
import re
import random
from appJar import gui
from time import sleep
import copy
from datetime import datetime
import pandas
import pyautogui
from pyautogui import press, typewrite, hotkey
import threading
import sys
import cv2


####################################################
########################################## ##########
def click2(button):  # MORAL SUPPORT AND TECHINCAL SUPPORT
    if button == "Technical Support":
        techsupport(introapp)
    elif button == "Moral Support":
        moral(introapp, list)


def support(introapp):  # DEFINES GUI WINDOW
    introapp.startSubWindow("Support")
    introapp.setGeometry("700x200")
    introapp.showSubWindow("Support")
    introapp.addButtons(["Technical Support"], click2, 0, 0)
    introapp.addButtons(["Moral Support"], click2, 0, 1)


####################################################
####################################################

def moral(introapp, list):  # ALL OF THE MORAL SUPPORT OPTIONS
    number = random.randint(0, 18)
    moraloptions = ["The difficulties in life are intended to make us better, not bitter",
                    "When you get to the end of your rope, tie a knot in it and hang on!",
                    "Fall seven times, stand up eight.",
                    "When life gives you a hundred reasons to cry, show life that you have a thousand reasons to smile.",
                    "Let your strongest muscle be the will.",
                    "When you get into a tight place and everything goes against you, till it seems as though you could not hang on a minute longer, never give up then, for that is just the place and time that the tide will turn.",
                    "Pain is inevitable, but suffering is optional.",
                    "I ask not for a lighter burden, but for broader shoulders.",
                    "Adversity has the effect of eliciting talents which, in prosperous circumstances, would have lain dormant.",
                    "Problems are only opportunities with thorns on them.",
                    "I have heard there are troubles of more than one kind." + "\n" + "Some come from ahead and some come from behind." + "\n" + "But I've bought a big bat. I'm all ready you see." + "\n" + "Now my troubles are going to have troubles with me!",
                    "Rock bottom is good solid ground, and a dead end street is just a place to turn around.",
                    "Count the garden by the flowers, never by the leaves that fall.  Count your life with smiles and not the tears that roll.",
                    "If one dream should fall and break into a thousand pieces, never be afraid to pick one of those pieces up and begin again.",
                    "The difference between perseverance and obstinacy is that one comes from a strong will, and the other from a strong won't.",
                    "Nobody trips over mountains. It is the small pebble that causes you to stumble. Pass all the pebbles in your path and you will find you have crossed the mountain.",
                    "I may not be there yet, but I'm closer than I was yesterday.",
                    "Problems are not stop signs, they are guidelines.",
                    "Look at a stone cutter hammering away at his rock, perhaps a hundred times without as much as a crack showing in it. Yet at the hundred-and-first blow it will split in two, and I know it was not the last blow that did it, but all that had gone before."]
    moralmessage = moraloptions[number]
    introapp.infoBox("Moral Support", moralmessage, parent=None)
    introapp.destroySubWindow("Support")


####################################################
####################################################

def techsupport(introapp):
    techmessage = "For technical assistance with this program, email Shreyas S B at Shreyas.SB@cerner.com"
    introapp.infoBox("Technical Support", techmessage, parent=None)
    introapp.destroySubWindow("Support")


####################################################
####################################################

def click1(button):
    infomessage = "CERNER CORPORATION" + "\n" + "User defined fields" + "\n" + "compiled: Jun 20 2021;  v1.0" + "\n" + "Copyright (C) 2021 Shreyas Coding Inc., Cerner Corp." + "\n" + "User Defined fields is free software for adding Custom defined fields and has no warranty whatsoever."
    helpmessage = "This program was written to build flex rules from a .xlsx file." + "\n" + "\n" + "To use this program, have the file saved as a .xlsx file.  To run the program, click the button 'RUN PROGRAM'. \n When the new window opens, click 'File' to open the file selection window.  Select the automation file and click 'Open'. \n Next, click 'Directory' and select the location you want the output log file to be saved.' \n  Finally, click 'RUN'."
    if button == "Program Info":
        introapp.infoBox("Information", infomessage, parent=None)
    elif button == "Program Help":
        introapp.infoBox("Help", helpmessage, parent=None)
    elif button == "Support":
        support(introapp)
    elif button == "CLOSE":
        introapp.stop()
        sys.exit()
    elif button == "RUN PROGRAM":
        introapp.stop()


introapp = gui("Main Screen", "1100x400")
introapp.setFont(size=22, family='Verdana')
introapp.setBg("white", override=False, tint=False)
introapp.addImage("pic", "cerner.png", 0, 0)
introapp.addLabel("title1", "User defined fields", 0, 1, 2)
introapp.addButtons(["Program Info"], click1, 1, 0)
introapp.addButtons(["Program Help"], click1, 1, 1)
introapp.addButtons(["Support"], click1, 1, 2)
introapp.addButtons(["RUN PROGRAM"], click1, 2, 0)
introapp.addButtons(["CLOSE"], click1, 2, 2)
introapp.go()
#################################################################################################################################

######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
######################################################################################################################################
###############################################################################################################################
def checkStop():
    return dcwapp.yesNoBox("Confirm Exit", "Are you sure you want to exit the application?")


dcwapp = gui("User defined fields", "600x400")
dcwapp.setBg("white", override=False, tint=False)
dcwapp.addLabel("title", "Welcome to User defined fields", 0, 0, 3)
dcwapp.setLabelBg("title", "white")
dcwapp.addFileEntry("Automation File", 1, 0, 3)
dcwapp.addDirectoryEntry("Log File Destination", 2, 0, 3)
dcwapp.setEntry("Automation File", "Enter the automation file here!")
dcwapp.setEntry("Log File Destination", "Where do you want the log file saved?")
dcwapp.setStopFunction(checkStop)


output_list = []

def tshoot():
    onscreen = None
    check = 0

    while onscreen == None and check < 3:
        check+=1
        onscreen = pyautogui.locateOnScreen('DuplicatedUniqueKey.png',confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('DuplicatedUniqueKey2.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('ErrorOcc.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('ErrorOcc1.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('ErrorOcc2.png', confidence=0.7)
        if onscreen != None:
            break

    if check >= 3:
        output_list.append('Success')

    else:
        output_list.append('Fail')
        identify1()
        identify2()
        identify3()


#t1 = 'Ok.png' or 'Ok2.png' or 'Ok1.png' or 'Ok3.png'
#t2 = 'Cancel.png' or 'Cancel1.png' or 'Cancel3.png' or 'Cancel4.png' or 'Cancel5.png'
#t3 = 'Yes.png' or 'Yes2.png' or 'Yes3.png'


def identify1():
    onscreen = None
    check = 0

    while onscreen == None and check < 3:
        check += 1
        onscreen = pyautogui.locateOnScreen('Ok.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Ok4.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Ok2.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Ok5.png', confidence=0.7)
        if onscreen != None:
            break
    if check >= 3:
        dcwapp.warningBox("Oops!", "I couldn't find the 'Ok' button.  Please click 'Ok' button and then close this message box")
        sleep(3)

    else:
        newfieldX, newfieldY = pyautogui.center(onscreen)
        pyautogui.moveTo(newfieldX, newfieldY)
        sleep(.05)
        if (newfieldX, newfieldY) != pyautogui.position():
            dcwapp.warningBox("Paused!",
                          "You moved the cursor! Program paused.  Place your cursor over the 'Ok' button.  After closing this message box, the program will resume in 5 seconds.")
            sleep(5)
        sleep(.05)
        pyautogui.click()
        #time.sleep(1)

def identify2():
    onscreen = None
    check = 0

    while onscreen == None and check < 3:
        check += 1
        onscreen = pyautogui.locateOnScreen('Cancel.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Cancel1.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Cancel3.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Cancel5.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Cancel4.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Cancel6.png', confidence=0.7)
        if onscreen != None:
            break

    if check >= 3:
        dcwapp.warningBox("Oops!", "I couldn't find the 'Cancel' button.  Please click the 'Cancel' button and then close this message box")
        sleep(3)

    else:
        newfieldX, newfieldY = pyautogui.center(onscreen)
        pyautogui.moveTo(newfieldX, newfieldY)
        sleep(.05)
        if (newfieldX, newfieldY) != pyautogui.position():
            dcwapp.warningBox("Paused!",
                          "You moved the cursor! Program paused.  Place your cursor over the 'Cancel' button.  After closing this message box, the program will resume in 5 seconds.")
            sleep(5)
        sleep(.05)
        pyautogui.click()
        #time.sleep(1)

def identify3():
    onscreen = None
    check = 0

    while onscreen == None and check < 3:
        check += 1
        onscreen = pyautogui.locateOnScreen('Yes.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Yes2.png', confidence=0.7)
        if onscreen != None:
            break
        onscreen = pyautogui.locateOnScreen('Yes3.png', confidence=0.7)
        if onscreen != None:
            break

    if check >= 3:
        dcwapp.warningBox('Oops',"Oops!", "I couldn't find the 'Yes' button.  Please Click 'Yes' button and then close this message box")
        sleep(3)

    else:
        newfieldX, newfieldY = pyautogui.center(onscreen)
        pyautogui.moveTo(newfieldX, newfieldY)
        sleep(.05)
        if (newfieldX, newfieldY) != pyautogui.position():
            dcwapp.warningBox("Paused!",
                          "You moved the cursor! Program paused.  Place your cursor over the 'Yes' button.  After closing this message box, the program will resume in 5 seconds.")
            sleep(5)
        sleep(.05)
        pyautogui.click()
        #time.sleep(1)



def identify(a):
    newfieldX, newfieldY = pyautogui.locateCenterOnScreen(a, confidence=0.7)
    pyautogui.moveTo(newfieldX, newfieldY)
    sleep(.05)
    if (newfieldX, newfieldY) != pyautogui.position():
        dcwapp.warningBox("Paused!",
                          "You moved the cursor! Program paused.  Place your cursor over the 'Add' button.  After closing this message box, the program will resume in 5 seconds.")
        sleep(5)
    sleep(.05)
    pyautogui.click()
    time.sleep(1)

def add(p):
    onscreen = None
    check = 0
    while onscreen == None and check < 3:
        onscreen = pyautogui.locateCenterOnScreen(p, confidence=0.7)
        if onscreen != None:
            break
        else:
            check += 1

    if check >= 3:
        dcwapp.warningBox('Oops!','Oops!, I could not find the Add button. Please click the button and then close this message box')
        sleep(3)

    else:
        identify(p)

def buttonThread():
    t1=threading.Thread(target=press1)
    t1.daemon= True
    t1.start()

def press1():
    try:
        myfile = str(dcwapp.getEntry("Automation File"))
        temp = str(dcwapp.getEntry("Log File Destination"))
        now = datetime.now()
        current_time = now.strftime("%H_%M_%S")
        filename = "LogFile"
        mynewfile = temp + "/" + filename + current_time + ".xlsx"
        excelfile = pandas.ExcelFile(myfile)
        workbook = xlsxwriter.Workbook(mynewfile)
        workbook.close()
        df = excelfile.parse(0)
        f1 = df['Field Name'].values.tolist()
        f2 = df['Unique Key'].values.tolist()
        f3 = df['PROMPT TYPE'].values.tolist()
        f4 = df['CODESET'].values.tolist()
        time.sleep(1)

        for name in range(len(f1)):  # typing appt name into the menmonic field in appt type tool to search for it
            time.sleep(1)
            p = 'Add5.png'
            add(p)
            pyautogui.write(f1[name])
            time.sleep(.5)
            keyboard.press('tab')
            time.sleep(.5)

            pyautogui.write(f2[name])
            time.sleep(.5)
            keyboard.press('tab')
            time.sleep(1)

            if f3[name] == 'Text' or f3[name] == 'TEXT':
                keyboard.press('tab')
                time.sleep(.5)
                keyboard.press('enter')
                time.sleep(.5)
                tshoot()

            elif f3[name] == 'Multi' or f3[name] == 'MULTI':
                i = 0
                while i < 1:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                keyboard.press('tab')
                time.sleep(1)
                keyboard.press('enter')
                time.sleep(1)
                tshoot()

            elif f3[name] == 'Date' or f3[name] == 'DATE':
                i = 0
                while i < 2:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                keyboard.press('tab')
                time.sleep(1)
                keyboard.press('enter')
                time.sleep(1)
                tshoot()

            elif f3[name] == 'Numeric' or f3[name] == 'NUMERIC' or f3[name] == "Number" or f3[name] == "NUMBER":
                i = 0
                while i < 3:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                keyboard.press('tab')
                time.sleep(1)
                keyboard.press('enter')
                time.sleep(1)
                tshoot()


            elif f3[name] == 'Coded' or f3[name] == 'CODED' or f3[name] == "Codified" or f3[name] == "CODIFIED":
                time.sleep(1)
                z = f4[name]
                i = 0
                while i < 4:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                keyboard.press('tab')
                time.sleep(1)
                pyautogui.write(str(z))
                keyboard.press('tab')
                time.sleep(1)
                keyboard.press('enter')
                time.sleep(1)
                tshoot()

            df1 = pd.DataFrame(list(zip(f1, f2, f3, f4, output_list)),
                               columns=['Field Name', 'Unique Key', 'PROMPT TYPE', 'CODESET',
                                        'Status'])
            df1.to_excel(mynewfile)
            """with pd.ExcelWriter(mynewfile, engine="openpyxl", mode='w') as writer:
                # with pd.ExcelWriter(desk_excel,engine= "openpyxl", mode='a') as writer:
                # df.to_excel(writer, index=False,sheet_name='Sheet1')
                df1.to_excel(writer, engine="openpyxl", sheet_name='Result')
                writer.save()"""

    except PermissionError:
        pyautogui.alert("Please close the excel workbook and then run automation.")
        exit()


def press(button):

    if button == "Close":
        dcwapp.stop()
        sys.exit()

    elif button == "Try it":
        myfile = str(dcwapp.getEntry("Automation File"))
        excelfile = pandas.ExcelFile(myfile)
        df = excelfile.parse(0)
        f1 = df['Field Name'].values.tolist()
        f2 = df['Unique Key'].values.tolist()
        f3 = df['PROMPT TYPE'].values.tolist()
        f4 = df['CODESET'].values.tolist()
        time.sleep(1)

        for name in range(len(f1)):
            if name>3:
                sys.exit()

            p = 'Add5.png'
            time.sleep(1)
            add(p)
            pyautogui.write(f1[name])
            time.sleep(.5)
            keyboard.press('tab')
            time.sleep(.5)

            pyautogui.write(f2[name])
            time.sleep(.5)
            keyboard.press('tab')
            time.sleep(1)

            if f3[name] == 'Text' or f3[name] == 'TEXT':
                #keyboard.press('tab')
                time.sleep(.5)
                identify2()
                identify3()

            elif f3[name] == 'Multi' or f3[name] == 'MULTI':
                i = 0
                while i < 1:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                #keyboard.press('tab')
                time.sleep(1)
                identify2()
                identify3()

            elif f3[name] == 'Date' or f3[name] == 'DATE':
                i = 0
                while i < 2:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                #keyboard.press('tab')
                time.sleep(1)
                identify2()
                identify3()

            elif f3[name] == 'Numeric' or f3[name] == 'NUMERIC' or f3[name] == "Number" or f3[name] == "NUMBER":
                i = 0
                while i < 3:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                #keyboard.press('tab')
                time.sleep(1)
                identify2()
                identify3()


            elif f3[name] == 'Coded' or f3[name] == 'CODED' or f3[name] == "Codified" or f3[name] == "CODIFIED":
                time.sleep(1)
                z = f4[name]
                i = 0
                while i < 4:
                    keyboard.press('down')
                    time.sleep(0.25)
                    i = i + 1
                keyboard.press('tab')
                time.sleep(1)
                pyautogui.write(str(z))
                #keyboard.press('tab')
                time.sleep(1)
                identify2()
                identify3()



dcwapp.addButtons(["Run"], buttonThread, 3, 0)
dcwapp.addButtons(["Try it"], press, 3, 1)
dcwapp.addButtons(["Close"], press, 3, 2)
dcwapp.go()

