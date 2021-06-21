from pyautogui import *
import pyautogui
import time
import keyboard
import pandas as pd
import xlsxwriter
import openpyxl
import sys, os #till here is SHREYAS imports
from operator import itemgetter  #######package to sort list of lists
import re
import random
from appJar import gui
from time import sleep
import copy
from datetime import datetime



"""try:
    from PIL import Image, ImageFilter, ImageEnhance
except ImportError:
    import Image, ImageFilter, ImageEnhance
#import pytesseract
#from pytesseract import Output"""
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
dcwapp = gui("User defined fields", "600x400")
dcwapp.setBg("white", override=False, tint=False)
dcwapp.addLabel("title", "Welcome to User defined fields", 0, 0, 2)
dcwapp.setLabelBg("title", "white")
dcwapp.addFileEntry("Automation File", 1, 0, 2)
dcwapp.addDirectoryEntry("Log File Destination", 2, 0, 2)
dcwapp.setEntry("Automation File", "Enter the automation file here!")
dcwapp.setEntry("Log File Destination", "Where do you want the log file saved?")

"""def exit_program():
    def on_press(key):
        if str(key) == 'Key.esc':
            main.status = 'pause'
            user_input = input('Program paused! What would you like to do? If you would like to Continue, Type 1. If you would like to skip the Userdefined field Type 2 ')
            while user_input != '1' and user_input != '2':
                user_input = input('Incorrect input, try either "y" or "n" ')

            if user_input == '1':
                main.status = 'run'
            elif user_input == 'n':
                main.status = 'exit'
                exit()

    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()"""


def press(button):
    if button == "CLOSE":
        dcwapp.stop()
        sys.exit()
    else:
        try:
            """main.status = 'run' 
            while True:
                #print('running')
                #time.sleep(1)

                while main.status == 'pause':
                    time.sleep(1)

                if main.status == 'exit':
                    print('Main program closing')
                    break""" #Could not figure out how to make use of this https://stackoverflow.com/questions/65068775/how-to-stop-or-pause-pyautogui-at-any-moment-that-i-want


            def resource_path(relative_path):
                if hasattr(sys, '_MEIPASS'):
                    return os.path.join(sys._MEIPASS, relative_path)
                return os.path.join(os.path.abspath("."), relative_path)

            """path1 = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Cerner Corporation', 'Desktop',
                                 'Automation Files', 'User Defined Fields', 'User_Defined_Fields.xlsx')
            path2 = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Cerner Corporation', 'Desktop',
                                 'Automation Files', 'User Defined Fields', 'Error Log.xlsx')"""

            myfile = str(dcwapp.getEntry("Automation File"))
            temp = str(dcwapp.getEntry("Log File Destination"))
            now = datetime.now()
            current_time = now.strftime("%H:%M:%S")

            filename = "LogFile"
            mynewfile = temp + "/" + filename + current_time + ".xlsx"

            excelfile = pandas.ExcelFile(myfile)
            df = excelfile.parse(0)
            #df = pd.read_excel(myfile, sheet_name="Sheet1")

            # df = pd.read_excel(r'C:\Users\SS078074\OneDrive - Cerner Corporation\Desktop\Automation Files\User Defined Fields\User_Defined_Fields.xlsx',sheet_name="Sheet1")
            #writer = pd.ExcelWriter(temp, engine="openpyxl", options={'strings_to_formulas': False})

            # appt_list = df['Appointment Mnemonics'].values.tolist()
            f1 = df['Field Name/Prompt Description'].values.tolist()
            f2 = df['CDF/Unique Key'].values.tolist()
            f3 = df['PROMPT_TYPE'].values.tolist()
            f4 = df['CODESET'].values.tolist()
            # declare list variable to hold the results of each appt name
            output_list = []
            time.sleep(1)
            a = pyautogui.locateCenterOnScreen(resource_path('Add3.png'), confidence=0.9)
            time.sleep(1)

            if a != None:
                x, y = pyautogui.locateCenterOnScreen(resource_path('Add3.png'), confidence=0.9)
                pyautogui.click(x, y)
                time.sleep(1)

                for name in range(len(f1)):  # typing appt name into the menmonic field in appt type tool to search for it
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

                        if (pyautogui.locateOnScreen(resource_path('DuplicatedUniqueKey1.png'),confidence=0.9) != None) or (
                                pyautogui.locateOnScreen(resource_path('ErrorOcc.png'), confidence=0.9) != None):

                            if (pyautogui.locateOnScreen(resource_path('Ok.png')) != None) or (pyautogui.locateOnScreen(resource_path('Ok2.png')) != None) or (
                                    pyautogui.locateOnScreen(resource_path('Ok1.png')) != None):
                                x, y = pyautogui.locateCenterOnScreen(
                                    resource_path('Ok.png') or resource_path('Ok2.png') or resource_path('Ok1.png'), confidence=0.9)
                                pyautogui.click(x, y)
                                time.sleep(4)

                                if pyautogui.locateOnScreen(resource_path('Cancel.png'), confidence=0.9) != None or pyautogui.locateOnScreen(resource_path('Cancel1.png'), confidence=0.9) != None:
                                    x, y = pyautogui.locateCenterOnScreen(resource_path('Cancel.png') or resource_path('Cancel1.png'), confidence=0.9)
                                    pyautogui.click(x, y)
                                    time.sleep(4)

                                    if pyautogui.locateOnScreen(resource_path('Yes.png')) != None or pyautogui.locateOnScreen(resource_path('Yes1.png')) != None:
                                        x, y = pyautogui.locateCenterOnScreen(resource_path('Yes.png') or resource_path('Yes1.png'), confidence=0.9)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        output_list.append('Fail')
                        else:
                            output_list.append('Success')

                    elif f3[name] == 'Multi' or f3[name] == 'MULTI':
                        i = 0
                        while i < 1:
                            keyboard.press('down')
                            time.sleep(0.25)
                            i = i + 1
                        keyboard.press('tab')
                        time.sleep(1)
                        keyboard.press('enter')
                        time.sleep(4)

                        if (pyautogui.locateOnScreen(resource_path('DuplicatedUniqueKey1.png'),
                                                     confidence=0.9) != None) or (
                                pyautogui.locateOnScreen(resource_path('ErrorOcc.png'), confidence=0.9) != None):

                            if (pyautogui.locateOnScreen(resource_path('Ok.png')) != None) or (pyautogui.locateOnScreen(resource_path('Ok2.png')) != None) or (
                                    pyautogui.locateOnScreen(resource_path('Ok1.png')) != None):
                                x, y = pyautogui.locateCenterOnScreen(
                                    resource_path('Ok.png') or resource_path('Ok2.png') or resource_path('Ok1.png'), confidence=0.9)
                                pyautogui.click(x, y)
                                time.sleep(4)

                                if pyautogui.locateOnScreen(resource_path('Cancel.png'), confidence=0.9) != None or pyautogui.locateOnScreen(resource_path('Cancel1.png'), confidence=0.9) != None:
                                    x, y = pyautogui.locateCenterOnScreen(resource_path('Cancel.png') or resource_path('Cancel1.png'), confidence=0.9)
                                    pyautogui.click(x, y)
                                    time.sleep(4)

                                    if pyautogui.locateOnScreen(resource_path('Yes.png')) != None or pyautogui.locateOnScreen(resource_path('Yes1.png')) != None:
                                        x, y = pyautogui.locateCenterOnScreen(resource_path('Yes.png') or resource_path('Yes1.png'), confidence=0.9)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        output_list.append('Fail')

                        else:
                            output_list.append('Success')

                    elif f3[name] == 'Date' or f3[name] == 'DATE':
                        i = 0
                        while i < 2:
                            keyboard.press('down')
                            time.sleep(0.25)
                            i = i + 1
                        keyboard.press('tab')
                        time.sleep(1)
                        keyboard.press('enter')
                        time.sleep(4)

                        if (pyautogui.locateOnScreen(resource_path('DuplicatedUniqueKey1.png'),
                                                     confidence=0.9) != None) or (
                                pyautogui.locateOnScreen(resource_path('ErrorOcc.png'), confidence=0.9) != None):

                            if (pyautogui.locateOnScreen(resource_path('Ok.png')) != None) or (pyautogui.locateOnScreen(resource_path('Ok2.png')) != None) or (
                                    pyautogui.locateOnScreen(resource_path('Ok1.png')) != None):
                                x, y = pyautogui.locateCenterOnScreen(
                                    resource_path('Ok.png') or resource_path('Ok2.png') or resource_path('Ok1.png'), confidence=0.9)
                                pyautogui.click(x, y)
                                time.sleep(4)

                                if pyautogui.locateOnScreen(resource_path('Cancel.png'), confidence=0.9) != None or pyautogui.locateOnScreen(resource_path('Cancel1.png'), confidence=0.9) != None:
                                    x, y = pyautogui.locateCenterOnScreen(resource_path('Cancel.png') or resource_path('Cancel1.png'), confidence=0.9)
                                    pyautogui.click(x, y)
                                    time.sleep(4)

                                    if pyautogui.locateOnScreen(resource_path('Yes.png')) != None or pyautogui.locateOnScreen(resource_path('Yes1.png')) != None:
                                        x, y = pyautogui.locateCenterOnScreen(resource_path('Yes.png') or resource_path('Yes1.png'), confidence=0.9)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        output_list.append('Fail')

                        else:
                            output_list.append('Success')

                    elif f3[name] == 'Numeric' or f3[name] == 'NUMERIC' or f3[name] == "Number" or f3[name] == "NUMBER":
                        i = 0
                        while i < 3:
                            keyboard.press('down')
                            time.sleep(0.25)
                            i = i + 1
                        keyboard.press('tab')
                        time.sleep(1)
                        keyboard.press('enter')
                        time.sleep(4)

                        if (pyautogui.locateOnScreen(resource_path('DuplicatedUniqueKey1.png'),
                                                     confidence=0.9) != None) or (
                                pyautogui.locateOnScreen(resource_path('ErrorOcc.png'), confidence=0.9) != None):

                            if (pyautogui.locateOnScreen(resource_path('Ok.png')) != None) or (pyautogui.locateOnScreen(resource_path('Ok2.png')) != None) or (
                                    pyautogui.locateOnScreen(resource_path('Ok1.png')) != None):
                                x, y = pyautogui.locateCenterOnScreen(
                                    resource_path('Ok.png') or resource_path('Ok2.png') or resource_path('Ok1.png'), confidence=0.9)
                                pyautogui.click(x, y)
                                time.sleep(4)

                                if pyautogui.locateOnScreen(resource_path('Cancel.png'), confidence=0.9) != None or pyautogui.locateOnScreen(resource_path('Cancel1.png'), confidence=0.9) != None:
                                    x, y = pyautogui.locateCenterOnScreen(resource_path('Cancel.png') or resource_path('Cancel1.png'), confidence=0.9)
                                    pyautogui.click(x, y)
                                    time.sleep(4)

                                    if pyautogui.locateOnScreen(resource_path('Yes.png')) != None or pyautogui.locateOnScreen(resource_path('Yes1.png')) != None:
                                        x, y = pyautogui.locateCenterOnScreen(resource_path('Yes.png') or resource_path('Yes1.png'), confidence=0.9)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        output_list.append('Fail')

                        else:
                            output_list.append('Success')

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
                        time.sleep(4)

                        if (pyautogui.locateOnScreen(resource_path('DuplicatedUniqueKey1.png'),
                                                     confidence=0.9) != None) or (
                                pyautogui.locateOnScreen(resource_path('ErrorOcc.png'), confidence=0.9) != None):

                            if (pyautogui.locateOnScreen(resource_path('Ok.png')) != None) or (pyautogui.locateOnScreen(resource_path('Ok2.png')) != None) or (
                                    pyautogui.locateOnScreen(resource_path('Ok1.png')) != None):
                                x, y = pyautogui.locateCenterOnScreen(resource_path('Ok.png') or resource_path('Ok2.png') or resource_path('Ok1.png'), confidence=0.9)
                                pyautogui.click(x, y)
                                time.sleep(4)

                                if pyautogui.locateOnScreen(resource_path('Cancel.png')) != None or pyautogui.locateOnScreen(resource_path('Cancel1.png')) != None:
                                    x, y = pyautogui.locateCenterOnScreen(resource_path('Cancel.png') or resource_path('Cancel1.png'), confidence=0.5)
                                    pyautogui.click(x, y)
                                    time.sleep(4)

                                    if pyautogui.locateOnScreen(resource_path('Yes.png')) != None or pyautogui.locateOnScreen(resource_path('Yes1.png')) != None:
                                        x, y = pyautogui.locateCenterOnScreen(resource_path('Yes.png') or resource_path('Yes1.png'), confidence=0.9)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        output_list.append('Fail')

                        else:
                            output_list.append('Success')

                    """a = pyautogui.locateCenterOnScreen(resource_path('Add3.png'), confidence=0.9)
                    time.sleep(1)

                    if a != None:
                        x, y = pyautogui.locateCenterOnScreen(resource_path('Add3.png'), confidence=0.9)
                        pyautogui.click(x, y)
                        time.sleep(1)"""

                #pyautogui.alert("Automation Complete!")

            df1 = pd.DataFrame(list(zip(f1, f2, f3, f4, output_list)),
                               columns=['Field Name/Prompt Description', 'CDF/Unique Key', 'PROMPT_TYPE', 'CODESET',
                                        'Status'])
            df1.to_excel(mynewfile)
            dcwapp.infoBox("Success", "Ta-da!  Finished!")

        except PermissionError:
            pyautogui.alert("Please close the excel workbook and then run automation.")
            exit()



dcwapp.addButtons(["RUN"], press, 4, 0)
dcwapp.addButtons(["CLOSE"], press, 4, 1)
dcwapp.go()