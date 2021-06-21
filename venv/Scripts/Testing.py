from operator import itemgetter  #######package to sort list of lists
import re
import os
import sys
import random
from appJar import gui
from time import sleep
import copy

try:
    from PIL import Image, ImageFilter, ImageEnhance
except ImportError:
    import Image, ImageFilter, ImageEnhance
#import pytesseract
#from pytesseract import Output
import pandas
import pyautogui
from pyautogui import press, typewrite, hotkey

#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


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
    techmessage = "For technical assistance with this program, email Chad Robertson at chad.m.robertson@cerner.com."
    introapp.infoBox("Technical Support", techmessage, parent=None)
    introapp.destroySubWindow("Support")


####################################################
####################################################

def click1(button):
    infomessage = "CERNER CORPORATION" + "\n" + "FlexRuleBuilder" + "\n" + "compiled: Aug  30 2019;  v1.0" + "\n" + "Copyright (C) 2018 Robertson Coding Inc., Cerner Corp." + "\n" + "FlexRuleBuilder is free software created for building flex rules, and has no warranty whatsoever."
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
introapp.addLabel("title1", "FlexRuleBuilder1.0", 0, 1, 2)
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
dcwapp = gui("Flexrulebuilder1.0", "600x400")
dcwapp.setBg("white", override=False, tint=False)
dcwapp.addLabel("title", "Welcome to FlexRuleBuilder1.0", 0, 0, 2)
dcwapp.setLabelBg("title", "white")
dcwapp.addFileEntry("Automation File", 1, 0, 2)
dcwapp.addDirectoryEntry("Log File Destination", 2, 0, 2)
dcwapp.setEntry("Automation File", "Enter the automation file here!")
dcwapp.setEntry("Log File Destination", "Where do you want the log file saved?")


def press(button):
    alreadyexists = False
    filetype = True
    direxists = True
    if button == "CLOSE":
        dcwapp.stop()
        sys.exit()
    else:
        myfile = str(dcwapp.getEntry("Automation File"))
        temp = str(dcwapp.getEntry("Log File Destination"))
        filename = "LogFile"
        mynewfile = temp + "/" + filename + ".xlsx"
        ext = str(os.path.splitext(myfile)[1])
        if ext != ".xlsx":
            filetype = True
        elif ext == ".xlsx":
            filetype = False
        excelfile = pandas.ExcelFile(myfile)

    #################################

    def deleteContent(pfile):
        f = open(pfile, 'w+')
        f.close()
        return 0

    if alreadyexists == False and filetype == False and direxists == True:  # This is where the panda starts
        df = excelfile.parse(0)  # saving the automation template into a dataframe
        df = df.fillna('')  # fills empty cells
        log = df.copy()  # makes copy of automation template to use as log file
        errors = []  # list of any errors
        mylist = []
        newmnemonic = True
        cancelled = False
        syntax = False
        duplicate = False
        done = False
        for i in range(len(df)):
            mylist.append(i)
            skip = False
            currentrole = ""
            cancelled = False
            row = df.iloc[i]  #######begin iterating over the DCW
            onscreen = None
            check = 0
            currentrole = row[1]
            if newmnemonic == True:
                newmnemonic = False
                while onscreen == None and check < 8:  # looking for mnemonic box
                    check += 1
                    onscreen = pyautogui.locateOnScreen('mnemonic.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('mnemonic1.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!", "I couldn't find the 'mnemonic' textbox.  Please type " + str(
                        row[1]) + " into the mnemonic box, then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the 'Mnemonic' text box.  After closing this message box, the program will resume in 5 seconds.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    typewrite(str(row[1]).rstrip('x\a0'), interval=.03)

                onscreen = None
                check = 0
                while onscreen == None and check < 8:  # looking for rule tab
                    onscreen = pyautogui.locateOnScreen('rule.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('rule1.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    check += 1
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the 'rule' tab. Please click the 'rule' tab then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the 'Rule' tab.  After closing this message box, the program will resume in 5 seconds.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()

            if str(row[4]) != "":
                onscreen = None
                check = 0
                if "(" in str(row[3]) and "(SN)" not in str(row[3]):
                    while onscreen == None and check < 8:  # Looking left parenthesis
                        onscreen = pyautogui.locateOnScreen('lpar.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('lpar1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the left parenthesis.  Please click once on the left parenthesis then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the left parenthesis.  After closing this message box, the program will resume in 5 seconds.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        pyautogui.click()
                elif ")" in str(row[3]) and "(SN)" not in str(row[3]):
                    while onscreen == None and check < 8:  # Looking right parenthesis
                        onscreen = pyautogui.locateOnScreen('rpar.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('rpar1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the right parenthesis.  Please click once on the right parenthesis then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the right parenthesis.  After closing this message box, the program will resume in 5 seconds.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        pyautogui.click()
                onscreen = None
                check = 0
                while onscreen == None and check < 8:  # Looking for operand window
                    onscreen = pyautogui.locateOnScreen('operand.png', confidence=0.95, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('operand1.png', confidence=0.95, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the 'Operand' window. Please click once on the operand detail " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX = onscreen[0] + onscreen[2]
                    newfieldY = onscreen[1] + onscreen[3]
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the operand window.  After closing this message box, the program will resume in 5 seconds.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    if "(SN)" not in str(row[3]):
                        typewrite(str(row[3].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                    else:
                        typewrite(str(row[3]).rstrip('x\a0'), interval=.03)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:  # looking for operand selection
                    onscreen = pyautogui.locateOnScreen('blue.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('blue1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the operand selection. Please double click on the operand selection " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the operand selection " + str(
                                              row[
                                                  3]) + " then close this message box; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    pyautogui.click()
                    sleep(.05)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('eventdetail11.png', confidence=0.9, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('eventdetail12.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the operand detail window. Please click ONCE on the operand detail " + str(
                                          row[4]) + " , then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY + 15)
                    sleep(.05)
                    if (newfieldX, newfieldY + 15) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over any value in the operand detail window.  After closing this message box, the program will resume in 5 seconds.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    typewrite(str(row[4]).rstrip('x\a0'), interval=.03)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('detailblue.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('detailblue1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the operand detail selection. Please double click on the operand detail selection " + str(
                                          row[4]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over any value in the operand detail selection " + str(
                                              row[
                                                  4]) + ".  After closing this message box, the program will resume in 5 seconds.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    pyautogui.click()
                    sleep(.05)

            else:
                onscreen = None
                check = 0
                if "(" in str(row[3]) and "(SN)" not in str(row[3]):
                    while onscreen == None and check < 8:  # Looking left parenthesis
                        onscreen = pyautogui.locateOnScreen('lpar.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('lpar1png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the left parenthesis.  Please click once on the left parenthesis then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the left parenthesis.  After closing this message box, the program will resume in 5 seconds.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        pyautogui.click()
                        sleep(.05)
                elif ")" in str(row[3]) and "(SN)" not in str(row[3]):
                    while onscreen == None and check < 8:  # Looking right parenthesis
                        onscreen = pyautogui.locateOnScreen('rpar.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('rpar1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the right parenthesis.  Please click once on the right parenthesis then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the right parenthesis.  After closing this message box, the program will resume in 5 seconds.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        pyautogui.click()
                        sleep(.05)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('operand.png', confidence=0.95, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('operand1.png', confidence=0.95, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the 'Operand' window. Please click once on the operand detail " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX = onscreen[0] + onscreen[2]
                    newfieldY = onscreen[1] + onscreen[3]
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the operand window.  After closing this message box, the program will resume in 5 seconds.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    if "(SN)" not in str(row[3]):
                        typewrite(str(row[3].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                    else:
                        typewrite(str(row[3]).rstrip('x\a0'), interval=.03)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:  # looking for operand selection
                    onscreen = pyautogui.locateOnScreen('blue.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('blue1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the operand selection. Please double click on the operand selection " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the operand selection " + str(
                                              row[
                                                  3]) + " then close this message box; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    pyautogui.click()
                    sleep(.05)

            onscreen = None
            check = 0
            while onscreen == None and check < 8:
                onscreen = pyautogui.locateOnScreen('equals.png', confidence=0.9, grayscale=True)
                check += 1
                if onscreen != None:
                    break
                onscreen = pyautogui.locateOnScreen('equals1.png', confidence=0.9, grayscale=True)
                if onscreen != None:
                    break
            if check >= 8:
                dcwapp.warningBox("Oops!",
                                  "I couldn't find the 'equals' operator. Please double click on the 'equals' operator then close this message box.")
            else:
                newfieldX, newfieldY = pyautogui.center(onscreen)
                pyautogui.moveTo(newfieldX, newfieldY)
                sleep(.05)
                if (newfieldX, newfieldY) != pyautogui.position():
                    dcwapp.warningBox("Paused!",
                                      "You moved the cursor! Program paused.  Place your cursor over the 'equals' operator then close this message box; the program will resume in 5 seconds after closing.")
                    sleep(5)
                sleep(.05)
                pyautogui.click()
                pyautogui.click()
                sleep(.05)

            if str(row[4]) != "":  # if it's event detail
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('datasource.png', confidence=0.7, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('datasource1.png', confidence=0.7, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the 'equals' operator. Please click once on the datasource " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX = onscreen[0] + onscreen[2]
                    newfieldY = onscreen[1] + onscreen[3]
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the datasource window then close this message box; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    if "(SN)" not in str(row[3]):
                        typewrite(str(row[3].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                    else:
                        typewrite(str(row[3]).rstrip('x\a0'), interval=.03)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('blue.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('blue1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the datasource selection. Please double click on the datasource selection " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the datasource selection " + str(
                                              row[3]) + "; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    pyautogui.click()
                    sleep(.05)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('eventdetail11.png', confidence=0.9, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('eventdetail12.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the datasource detail selection. Please click once on the datasource detail selection " + str(
                                          row[4]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the datasource detail selection window; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    typewrite(str(row[4]).rstrip('x\a0'), interval=.03)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('detailblue.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('detailblue1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the datasource detail selection. Please double click on the datasource detail selection " + str(
                                          row[4]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the datasource detail selection " + str(
                                              row[4]) + "; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    pyautogui.click()
                    sleep(.05)
                #######Checking type of detail for yellow box, drop down, or ellipses
                box = False
                drop = False
                ellipse = False
                onscreen = None
                check = 0
                while onscreen == None and check < 8:  # testing for yellow box
                    onscreen = pyautogui.locateOnScreen('yellowbox.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('yellowbox1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if onscreen != None:
                    box = True
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the data entry box; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    if "(SN)" not in str(row[6]):
                        typewrite(str(row[6].rstrip("'".replace("(", "").replace(")", ""))).rstrip('x\a0'),
                                  interval=.03)
                    else:
                        typewrite(str(row[6]).rstrip('x\a0'), interval=.03)
                    pyautogui.press('enter')
                    sleep(.1)  # wait for surgeon validation
                onscreen = None
                check = 0
                while onscreen == None and check < 8 and box == False:  # testing for dropdown
                    onscreen = pyautogui.locateOnScreen('dropdown.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('dropdown1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if onscreen != None and box == False:
                    drop = True
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the data dropdown; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    dcwapp.warningBox("User Input Required", "Please choose option " + str(
                        row[6].replace("(", "").replace(")",
                                                        "")) + " from the dropdown.  After that, please close this diaglogue box.")
                onscreen = None
                ellipse = False
                check = 0
                while onscreen == None and check < 8 and drop == False:  # testing for ellipses
                    onscreen = pyautogui.locateOnScreen('ellipsesbox.png', confidence=0.9, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('ellipsesbox1.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8 and drop == False and box == False:
                    dcwapp.warningBox("Oops!", "I couldn't find data entry location. Please enter: " + str(
                        row[6].replace("(", "").replace(")",
                                                        "")) + " into the tool, click ok, then close this message box.")
                    skip = True
                else:
                    if onscreen != None:
                        ellipse = True
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the data entry box; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                        if "(SN)" not in str(row[6]):
                            typewrite(str(row[6].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                        else:
                            typewrite(str(row[6]).rstrip('x\a0'), interval=.03)
                        pyautogui.press('enter')
                    onscreen = None
                    check = 0
                    while onscreen == None and check < 8 and ellipse == True:  # clicking on ellipses
                        onscreen = pyautogui.locateOnScreen('ellipses.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('ellipses1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the ellipses button. Please click the ellipses button then close this message box.")
                    else:
                        if onscreen != None:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the ellipses button; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                    onscreen = None
                    check = 0
                    while onscreen == None and check < 8 and ellipse == True:  # clicking on value
                        onscreen = pyautogui.locateOnScreen('datavalue.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('datavalue1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!", "I couldn't find the data value. Please click the data value " + str(
                            row[6].replace("(", "").replace(")", "")) + " then close this message box.")
                    else:
                        if onscreen != None:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX + 15, newfieldY + 15)
                            sleep(.05)
                            if (newfieldX + 15, newfieldY + 15) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over data value " + str(
                                                      row[6].replace("(", "").replace(")",
                                                                                      "")) + "; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                    if ellipse == True:
                        onscreen = None  # pressing OK button
                        while onscreen == None and check < 8:
                            onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                            check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('ok1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                        if check >= 8:
                            dcwapp.warningBox("Oops!",
                                              "I couldn't find the ok button. Please click the ok button then close this message box.")
                        else:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                    onscreen = None  # pressing OK button for non ellipses
                    check = 0
                    while onscreen == None and ellipse == False and check < 8:
                        onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('ok1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the ok button. Please click the ok button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                onscreen = None
                validation = False
                check = 0
                while onscreen == None and check < 8:  # testing for errors
                    onscreen = pyautogui.locateOnScreen('validation.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('validation1.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('validation2.png', confidence=0.9, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('validation3.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                if onscreen != None:
                    validation = True
                    errors.append("Datasource detail could not be validated")
                    log['Errors'] = pandas.Series(errors)
                    log.to_excel(mynewfile)
                    onscreen = None  # pressing OK button
                    onscreen1 = None
                    check = 0
                    check1 = 0
                    while onscreen == None and check < 8:
                        onscreen = pyautogui.locate("ok.png", "haystack.png", confidence=.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locate('ok1.png', "haystack1.png", confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    while onscreen1 == None and check1 < 8:
                        onscreen1 = pyautogui.locateOnScreen("haystack.png", confidence=.9, grayscale=True)
                        check1 += 1
                        if onscreen1 != None:
                            break
                        onscreen1 = pyautogui.locateOnScreen('haystack1.png', confidence=0.9, grayscale=True)
                        if onscreen1 != None:
                            break
                    if check >= 8 or check1 >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the ok button. Please click the ok button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        newfieldX1 = onscreen1[0]
                        newfieldY1 = onscreen1[1]
                        pyautogui.moveTo(newfieldX + newfieldX1, newfieldY + newfieldY1)
                        sleep(.05)
                        if (newfieldX + newfieldX1, newfieldY + newfieldY1) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                    onscreen = None  # pressing cancelling out of yellow data box
                    check = 0
                    while onscreen == None and check < 8:
                        print("test1")
                        onscreen = pyautogui.locateOnScreen('cancel.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('cancel1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the cancel button. Please click the cancel button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the cancel button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                    onscreen = None  # clearing out entry
                    check = 0
                    while onscreen == None and check < 8:
                        onscreen = pyautogui.locateOnScreen('clear.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('clear1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the clear button. Please click the clear button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the clear button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(1)
                        pyautogui.click()
                        sleep(.05)
                    onscreen = None  # confirming clear
                    check = 0
                    while onscreen == None and check < 8:
                        onscreen = pyautogui.locateOnScreen('yes.png', confidence=0.75, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('yes1.png', confidence=0.75, grayscale=True)
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('yes2.png', confidence=0.75, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the yes button. Please click the yes button then close this message box.")
                        cancelled = True
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the yes button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                        cancelled = True
                    newmnemonic = True

            else:  # not event detail
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('datasource.png', confidence=0.7, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('datasource1.png', confidence=0.7, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the 'equals' operator. Please click once on the datasource " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX = onscreen[0] + onscreen[2]
                    newfieldY = onscreen[1] + onscreen[3]
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the datasource window then close this message box; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    if "(SN)" not in str(row[3]):
                        typewrite(str(row[3].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                    else:
                        typewrite(str(row[3]).rstrip('x\a0'), interval=.03)
                onscreen = None
                check = 0
                while onscreen == None and check < 8:
                    onscreen = pyautogui.locateOnScreen('blue.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('blue1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if check >= 8:
                    dcwapp.warningBox("Oops!",
                                      "I couldn't find the datasource selection. Please double click on the datasource selection " + str(
                                          row[3]) + " then close this message box.")
                else:
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the datasource selection " + str(
                                              row[3]) + "; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    pyautogui.click()
                    sleep(.05)

                onscreen = None
                box = False
                drop = False
                ellipse = False
                check = 0
                while onscreen == None and check < 8:  # testing for yellow box
                    onscreen = pyautogui.locateOnScreen('yellowbox.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('yellowbox1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if onscreen != None:
                    box = True
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the data entry box; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    pyautogui.click()
                    sleep(.05)
                    if "(SN)" not in str(row[6]):
                        typewrite(str(row[6].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                    else:
                        typewrite(str(row[6]).rstrip('x\a0'), interval=.03)
                    pyautogui.press('enter')
                    sleep(.1)  # wait for surgeon validation
                onscreen = None
                check = 0
                while onscreen == None and check < 8 and box == False:  # testing for dropdown
                    onscreen = pyautogui.locateOnScreen('dropdown.png', confidence=0.9, grayscale=False)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('dropdown1.png', confidence=0.9, grayscale=False)
                    if onscreen != None:
                        break
                if onscreen != None and box == False:
                    drop = True
                    newfieldX, newfieldY = pyautogui.center(onscreen)
                    pyautogui.moveTo(newfieldX, newfieldY)
                    sleep(.05)
                    if (newfieldX, newfieldY) != pyautogui.position():
                        dcwapp.warningBox("Paused!",
                                          "You moved the cursor! Program paused.  Place your cursor over the data dropdown; the program will resume in 5 seconds after closing.")
                        sleep(5)
                    sleep(.05)
                    dcwapp.warningBox("User Input Required", "Please choose option " + str(
                        row[6].replace("(", "").replace(")",
                                                        "")) + " from the dropdown.  After that, please close this diaglogue box.")
                onscreen = None
                ellipse = False
                check = 0
                while onscreen == None and check < 8 and drop == False:  # testing for ellipses
                    onscreen = pyautogui.locateOnScreen('ellipsesbox.png', confidence=0.9, grayscale=True)
                    check += 1
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('ellipsesbox1.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                if check >= 8 and drop == False and box == False:
                    dcwapp.warningBox("Oops!", "I couldn't find data entry location. Please enter: " + str(
                        row[6].replace("(", "").replace(")",
                                                        "")) + " into the tool, click ok, then close this message box.")
                    skip = True
                else:
                    if onscreen != None:
                        ellipse = True
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the data entry box; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                        if "(SN)" not in str(row[6]):
                            typewrite(str(row[6].replace("(", "").replace(")", "")).rstrip('x\a0'), interval=.03)
                        else:
                            typewrite(str(row[6]).rstrip('x\a0'), interval=.03)
                        pyautogui.press('enter')
                    onscreen = None
                    check = 0
                    while onscreen == None and check < 8:  # clicking on ellipses
                        onscreen = pyautogui.locateOnScreen('ellipses.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('ellipses1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the ellipses button. Please click the ellipses button then close this message box.")
                    else:
                        if onscreen != None:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the ellipses button; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                    onscreen = None
                    check = 0
                    while onscreen == None and check < 8:  # clicking on value
                        onscreen = pyautogui.locateOnScreen('datavalue.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('datavalue1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!", "I couldn't find the data value. Please click the data value " + str(
                            row[6].replace("(", "").replace(")", "")) + " then close this message box.")
                    else:
                        if onscreen != None:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX + 15, newfieldY + 15)
                            sleep(.05)
                            if (newfieldX + 15, newfieldY + 15) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over data value " + str(
                                                      row[6].replace("(", "").replace(")",
                                                                                      "")) + "; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                    if skip == False:
                        onscreen = None  # pressing OK button
                        while onscreen == None and check < 8:
                            onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('ok1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if check >= 8:
                            dcwapp.warningBox("Oops!",
                                              "I couldn't find the ok button. Please click the ok button then close this message box.")
                        else:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                onscreen = None
                validation = False
                check = 0
                while onscreen == None and check < 8:  # testing for errors
                    onscreen = pyautogui.locateOnScreen('validation.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('validation1.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('validation2.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    onscreen = pyautogui.locateOnScreen('validation3.png', confidence=0.9, grayscale=True)
                    if onscreen != None:
                        break
                    check += 1
                if onscreen != None:
                    validation = True
                    errors.append("Datasource detail could not be validated")
                    log['Errors'] = pandas.Series(errors)
                    log.to_excel(mynewfile)
                    onscreen = None  # pressing OK button
                    check = 0
                    while onscreen == None and check < 8:
                        onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the ok button. Please click the ok button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                    onscreen = None  # pressing cancelling out of yellow data box
                    check = 0
                    while onscreen == None and check < 8:
                        print("test2")
                        onscreen = pyautogui.locateOnScreen('cancel.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('cancel1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the cancel button. Please click the cancel button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the cancel button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                    onscreen = None  # clearing out entry
                    check = 0
                    while onscreen == None and check < 8:
                        onscreen = pyautogui.locateOnScreen('clear.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('clear1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the clear button. Please click the clear button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the clear button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                    onscreen = None  # confirming clear
                    check = 0
                    while onscreen == None and check < 8:
                        onscreen = pyautogui.locateOnScreen('yes.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('yes1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the yes button. Please click the yes button then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the yes button; the program will resume in 5 seconds after closing.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        sleep(.05)
                        cancelled = True
                    newmnemonic = True

            if cancelled == False:
                onscreen = None
                check = 0
                if "(" in str(row[6]) and "(SN)" not in str(row[6]):
                    while onscreen == None and check < 8:  # Looking left parenthesis
                        onscreen = pyautogui.locateOnScreen('lpar.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('lpar1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the left parenthesis.  Please click once on the left parenthesis then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the left parenthesis.  After closing this message box, the program will resume in 5 seconds.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        pyautogui.click()
                        sleep(.05)
                elif ")" in str(row[6]) and "(SN)" not in str(row[6]):
                    while onscreen == None and check < 8:  # Looking right parenthesis
                        onscreen = pyautogui.locateOnScreen('rpar.png', confidence=0.9, grayscale=True)
                        check += 1
                        if onscreen != None:
                            break
                        onscreen = pyautogui.locateOnScreen('rpar1.png', confidence=0.9, grayscale=True)
                        if onscreen != None:
                            break
                    if check >= 8:
                        dcwapp.warningBox("Oops!",
                                          "I couldn't find the right parenthesis.  Please click once on the right parenthesis then close this message box.")
                    else:
                        newfieldX, newfieldY = pyautogui.center(onscreen)
                        pyautogui.moveTo(newfieldX, newfieldY)
                        sleep(.05)
                        if (newfieldX, newfieldY) != pyautogui.position():
                            dcwapp.warningBox("Paused!",
                                              "You moved the cursor! Program paused.  Place your cursor over the right parenthesis.  After closing this message box, the program will resume in 5 seconds.")
                            sleep(5)
                        sleep(.05)
                        pyautogui.click()
                        pyautogui.click()
                        sleep(.05)
                if str(row[7]) != "":
                    if str(row[7]).upper() == "AND":
                        onscreen = None
                        check = 0
                        while onscreen == None and check < 8:  # Looking for and
                            onscreen = pyautogui.locateOnScreen('and.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('and1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if check >= 8:
                            dcwapp.warningBox("Oops!",
                                              "I couldn't find the and option.  Please click twice on and then close this message box.")
                        else:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the and option.  After closing this message box, the program will resume in 5 seconds.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            pyautogui.click()
                            sleep(.05)
                    elif str(row[7]).upper() == "OR":
                        onscreen = None
                        check = 0
                        while onscreen == None and check < 8:  # Looking for or
                            onscreen = pyautogui.locateOnScreen('or.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('or1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if check >= 8:
                            dcwapp.warningBox("Oops!",
                                              "I couldn't find the or.  Please click twice on or then close this message box.")
                        else:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the or.  After closing this message box, the program will resume in 5 seconds.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            pyautogui.click()
                            sleep(.05)

                try:
                    if df.iloc[i + 1][1] != currentrole:
                        newmnemonic = True
                        onscreen = None
                        check = 0
                        while onscreen == None and check < 8:
                            onscreen = pyautogui.locateOnScreen('save.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('save1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if check >= 8:
                            dcwapp.warningBox("Oops!",
                                              "I couldn't find the save button. Please click the save button then close this message box.")
                        else:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the save button; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                        onscreen = None
                        syntax = False
                        duplicate = False
                        check = 0
                        while onscreen == None and check < 8:  # testing for duplicate error
                            onscreen = pyautogui.locateOnScreen('duplicate.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('duplicate1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if onscreen != None:
                            duplicate = True
                            cancelled = True
                            errors.append("Mnemonic already exists in domain")
                            log['Errors'] = pandas.Series(errors)
                            log.to_excel(mynewfile)
                            for number in mylist:
                                errors.append("Mnemonic already exists in domain")
                                log['Errors'] = pandas.Series(errors)
                                log.to_excel(mynewfile)
                            mylist = []
                        onscreen == None
                        check = 0
                        while onscreen == None and check < 8 and duplicate == False:  # testing for syntax error
                            onscreen = pyautogui.locateOnScreen('syntax.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('syntax1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if onscreen != None:
                            syntax = True
                            cancelled = True
                            if syntax == True and duplicate == False:
                                errors.append("No datasource value entered, syntax error")
                                log['Errors'] = pandas.Series(errors)
                                log.to_excel(mynewfile)
                                for number in mylist:
                                    errors.append("No datasource value entered, syntax error")
                                    log['Errors'] = pandas.Series(errors)
                                    log.to_excel(mynewfile)
                                mylist = []
                            onscreen = None  # closing error message
                            check = 0
                            while onscreen == None and check < 8:
                                onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                                check += 1
                                if onscreen != None:
                                    break
                                onscreen = pyautogui.locateOnScreen('ok1.png', confidence=0.9, grayscale=True)
                                if onscreen != None:
                                    break
                            if check >= 8:
                                dcwapp.warningBox("Oops!",
                                                  "I couldn't find the ok button. Please click the ok button then close this message box.")
                            else:
                                newfieldX, newfieldY = pyautogui.center(onscreen)
                                pyautogui.moveTo(newfieldX, newfieldY)
                                sleep(.05)
                                if (newfieldX, newfieldY) != pyautogui.position():
                                    dcwapp.warningBox("Paused!",
                                                      "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                                    sleep(5)
                                sleep(.05)
                                pyautogui.click()
                                sleep(.05)
                            onscreen = None  # clearing out entry
                            check = 0
                            while onscreen == None and check < 8:
                                onscreen = pyautogui.locateOnScreen('clear.png', confidence=0.9, grayscale=True)
                                check += 1
                                if onscreen != None:
                                    break
                                onscreen = pyautogui.locateOnScreen('clear1.png', confidence=0.9, grayscale=True)
                                if onscreen != None:
                                    break
                            if check >= 8:
                                dcwapp.warningBox("Oops!",
                                                  "I couldn't find the clear button. Please click the clear button then close this message box.")
                            else:
                                newfieldX, newfieldY = pyautogui.center(onscreen)
                                pyautogui.moveTo(newfieldX, newfieldY)
                                sleep(.05)
                                if (newfieldX, newfieldY) != pyautogui.position():
                                    dcwapp.warningBox("Paused!",
                                                      "You moved the cursor! Program paused.  Place your cursor over the clear button; the program will resume in 5 seconds after closing.")
                                    sleep(5)
                                sleep(.05)
                                pyautogui.click()
                                sleep(.05)
                            onscreen = None  # confirming clear
                            check = 0
                            while onscreen == None and check < 8:
                                onscreen = pyautogui.locateOnScreen('yes.png', confidence=0.9, grayscale=True)
                                check += 1
                                if onscreen != None:
                                    break
                                onscreen = pyautogui.locateOnScreen('yes1.png', confidence=0.9, grayscale=True)
                                if onscreen != None:
                                    break
                            if check >= 8:
                                dcwapp.warningBox("Oops!",
                                                  "I couldn't find the yes button. Please click the yes button then close this message box.")
                            else:
                                newfieldX, newfieldY = pyautogui.center(onscreen)
                                pyautogui.moveTo(newfieldX, newfieldY)
                                sleep(.05)
                                if (newfieldX, newfieldY) != pyautogui.position():
                                    dcwapp.warningBox("Paused!",
                                                      "You moved the cursor! Program paused.  Place your cursor over the yes button; the program will resume in 5 seconds after closing.")
                                    sleep(5)
                                sleep(.05)
                                pyautogui.click()
                                sleep(.05)
                            cancelled = True
                            newmnemonic = True
                        if cancelled == False:
                            errors.append("Row built successfully")
                            log['Errors'] = pandas.Series(errors)
                            log.to_excel(mynewfile)
                            for number in mylist:
                                errors.append("Row built successfully")
                                log['Errors'] = pandas.Series(errors)
                                log.to_excel(mynewfile)
                            mylist = []



                except IndexError:
                    done = True
                    pass
                    if done == True:
                        onscreen = None
                        check = 0
                        while onscreen == None and check < 8:
                            onscreen = pyautogui.locateOnScreen('save.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('save1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if check >= 8:
                            dcwapp.warningBox("Oops!",
                                              "I couldn't find the save button. Please click the save button then close this message box.")
                        else:
                            newfieldX, newfieldY = pyautogui.center(onscreen)
                            pyautogui.moveTo(newfieldX, newfieldY)
                            sleep(.05)
                            if (newfieldX, newfieldY) != pyautogui.position():
                                dcwapp.warningBox("Paused!",
                                                  "You moved the cursor! Program paused.  Place your cursor over the save button; the program will resume in 5 seconds after closing.")
                                sleep(5)
                            sleep(.05)
                            pyautogui.click()
                            sleep(.05)
                        onscreen = None
                        syntax = False
                        duplicate = False
                        check = 0
                        while onscreen == None and check < 8:  # testing for duplicate error
                            onscreen = pyautogui.locateOnScreen('duplicate.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('duplicate1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if onscreen != None:
                            duplicate = True
                            cancelled = True
                            errors.append("Mnemonic already exists in domain")
                            log['Errors'] = pandas.Series(errors)
                            log.to_excel(mynewfile)
                            for number in mylist:
                                errors.append("Mnemonic already exists in domain")
                                log['Errors'] = pandas.Series(errors)
                                log.to_excel(mynewfile)
                            mylist = []
                        onscreen == None
                        check = 0
                        while onscreen == None and check < 8 and duplicate == False:  # testing for syntax error
                            onscreen = pyautogui.locateOnScreen('syntax.png', confidence=0.9, grayscale=True)
                            check += 1
                            if onscreen != None:
                                break
                            onscreen = pyautogui.locateOnScreen('syntax1.png', confidence=0.9, grayscale=True)
                            if onscreen != None:
                                break
                        if onscreen != None:
                            syntax = True
                            cancelled = True
                            if syntax == True and duplicate == False:
                                errors.append("No datasource value entered, syntax error")
                                log['Errors'] = pandas.Series(errors)
                                log.to_excel(mynewfile)
                                for number in mylist:
                                    errors.append("Mnemonic already exists in domain")
                                    log['Errors'] = pandas.Series(errors)
                                    log.to_excel(mynewfile)
                                mylist = []
                            onscreen = None  # closing error message
                            check = 0
                            while onscreen == None and check < 8:
                                onscreen = pyautogui.locateOnScreen('ok.png', confidence=0.9, grayscale=True)
                                check += 1
                                if onscreen != None:
                                    break
                                onscreen = pyautogui.locateOnScreen('ok1.png', confidence=0.9, grayscale=True)
                                if onscreen != None:
                                    break
                            if check >= 8:
                                dcwapp.warningBox("Oops!",
                                                  "I couldn't find the ok button. Please click the ok button then close this message box.")
                            else:
                                newfieldX, newfieldY = pyautogui.center(onscreen)
                                pyautogui.moveTo(newfieldX, newfieldY)
                                sleep(.05)
                                if (newfieldX, newfieldY) != pyautogui.position():
                                    dcwapp.warningBox("Paused!",
                                                      "You moved the cursor! Program paused.  Place your cursor over the ok button; the program will resume in 5 seconds after closing.")
                                    sleep(5)
                                sleep(.05)
                                pyautogui.click()
                                sleep(.05)
                            onscreen = None  # clearing out entry
                            check = 0
                            while onscreen == None and check < 8:
                                onscreen = pyautogui.locateOnScreen('clear.png', confidence=0.9, grayscale=True)
                                check += 1
                                if onscreen != None:
                                    break
                                onscreen = pyautogui.locateOnScreen('clear1.png', confidence=0.9, grayscale=True)
                                if onscreen != None:
                                    break
                            if check >= 8:
                                dcwapp.warningBox("Oops!",
                                                  "I couldn't find the clear button. Please click the clear button then close this message box.")
                            else:
                                newfieldX, newfieldY = pyautogui.center(onscreen)
                                pyautogui.moveTo(newfieldX, newfieldY)
                                sleep(.05)
                                if (newfieldX, newfieldY) != pyautogui.position():
                                    dcwapp.warningBox("Paused!",
                                                      "You moved the cursor! Program paused.  Place your cursor over the clear button; the program will resume in 5 seconds after closing.")
                                    sleep(5)
                                sleep(.05)
                                pyautogui.click()
                                sleep(.05)
                            onscreen = None  # confirming clear
                            check = 0
                            while onscreen == None and check < 8:
                                onscreen = pyautogui.locateOnScreen('yes.png', confidence=0.9, grayscale=True)
                                check += 1
                                if onscreen != None:
                                    break
                                onscreen = pyautogui.locateOnScreen('yes1.png', confidence=0.9, grayscale=True)
                                if onscreen != None:
                                    break
                            if check >= 8:
                                dcwapp.warningBox("Oops!",
                                                  "I couldn't find the yes button. Please click the yes button then close this message box.")
                            else:
                                newfieldX, newfieldY = pyautogui.center(onscreen)
                                pyautogui.moveTo(newfieldX, newfieldY)
                                sleep(.05)
                                if (newfieldX, newfieldY) != pyautogui.position():
                                    dcwapp.warningBox("Paused!",
                                                      "You moved the cursor! Program paused.  Place your cursor over the yes button; the program will resume in 5 seconds after closing.")
                                    sleep(5)
                                sleep(.05)
                                pyautogui.click()
                                sleep(.05)
                            cancelled = True
                            newmnemonic = True
                        if cancelled == False:
                            errors.append("Row built successfully")
                            log['Errors'] = pandas.Series(errors)
                            log.to_excel(mynewfile)
                            for number in mylist:
                                errors.append("Row built successfully")
                                log['Errors'] = pandas.Series(errors)
                                log.to_excel(mynewfile)
                            mylist = []

        log['Errors'] = pandas.Series(errors)
        log.to_excel(mynewfile)
        dcwapp.infoBox("Success", "Ta-da!  Finished!")

    elif alreadyexists == True and filetype == False and direxists == True:
        dcwapp.warningBox("Error", "The output file name already exists.  Please choose another file name.",
                          parent=None)
        dcwapp.setEntryInvalid("Enter the Output File Name Here")

    elif alreadyexists == True and filetype == True and direxists == True:
        dcwapp.warningBox("Error", "The output file name already exists and you didn't choose a .xlsx file.",
                          parent=None)
        dcwapp.setEntryInvalid("Enter the Output File Name Here")

    elif alreadyexists == True and filetype == True and direxists == False:
        dcwapp.warningBox("Error",
                          "The output file name already exists, you didn't choose a .xlsx file, and the directory you chose does not exist.",
                          parent=None)
        dcwapp.setEntryInvalid("Enter the Output File Name Here")

    elif alreadyexists == False and filetype == True and direxists == True:
        dcwapp.warningBox("Error", "You didn't choose a .xlsx file.", parent=None)
        dcwapp.setEntryValid("Enter the Output File Name Here")

    elif alreadyexists == False and filetype == True and direxists == False:
        dcwapp.warningBox("Error", "You didn't choose a .xlsx file and the directory you chose does not exist.",
                          parent=None)
        dcwapp.setEntryValid("Enter the Output File Name Here")

    elif alreadyexists == False and filetype == False and direxists == False:
        dcwapp.warningBox("Error", "The directory you chose does not exist.", parent=None)
        dcwapp.setEntryValid("Enter the Output File Name Here")


dcwapp.addButtons(["RUN"], press, 4, 0)
dcwapp.addButtons(["CLOSE"], press, 4, 1)
dcwapp.go()