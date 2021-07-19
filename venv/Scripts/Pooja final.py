#this version allows user input for which single scheduling comment needs to be added for all the appts. This is the final one.

from pyautogui import *
import pyautogui
import time
import keyboard
import pandas as pd
import os
import xlsxwriter
from tkinter import *
from tkinter import messagebox
import tkinter.font as font
import threading
import sys

root = Tk()

#code to find user's desktop path and create new sched_comment.xlsx file on their desktop
def create_excel():

    usr = os.environ['USERPROFILE'] #
    desk_excel = os.path.join((os.environ['USERPROFILE']),'OneDrive - Cerner Corporation','Desktop','Sched_comments.xlsx')
    workbook = xlsxwriter.Workbook(desk_excel)
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Appointment Mnemonics')
    workbook.close()
    time.sleep(1)
    pyautogui.alert("Success! Excel template has been downloaded on your desktop.")

#code for displaying instrcutuions
def my_instruct():
    global label2
    top =Toplevel()
    label2 = Label(top, text="Scheduling Comments Automation Instructions")
    label2['font']=myFont
    label2.pack(padx=2, pady =2)
    v = Scrollbar(top)
    v.pack(side = RIGHT,fill=Y)

    t= Text(top,wrap=WORD,padx=50,pady=50)
    t.insert(END,"""Steps:\n\n1. Ensure that you have opened schtools.exe -> 'Appointment types' tool.\n2. Make sure the 'Appointment Types' tool screen is opened full screen and there is no other window obstructing the word 'Mnemonic:' in the tool window.\n3. Then click on the "Download Excel Template" button found on this automation tool window. An excel sheet 'Sched_Comments' is downloaded on your desktop.\n4. Please enter the appt mnemonics (which need scheduling comments associated) under the column 'Appointment Mnemonics' in the excel file.\n5. Save and close the excel.\n6. Click on the "Run Automation" button found on this automation tool window.\n7. A pop-up will ask you to specify the Scheduling Comment mnemonic. Please enter comment mnemonic exactly as is in domain with correct capitalization, spaces, etc. and click OK.\n8. Sit back and watch the program complete the association in front of you.\n9. Once automation is done you will receive "Automation Complete!" pop-up.\n10. The 'Sched_Comments' excel sheet can now be opened and the results of each appt mnemonic will be given on the 'Results' tab of the sheet under the 'Status' column.\n\nIMP NOTE!!!\n\nIf you need to stop the program during its exceution, due to any reason:\n1. please open the main automation tool window and click 'Quit' button.\n2. The automation will stop immediately.\n3. All appointments upto the point that you quit, will be saved on the excel sheet 'Sched_Comments' under 'Results' tab.\n4. However the program will NOT restart from where you left off. You will need to update the excel sheet by removing the appt names which had completed upto the moment you pressed 'Quit' and then restart the automation,i.e., run the .exe file and then click on 'Run Automation' in the automation tool window.\n\nRemember - The 'Results' sheet will capture the results upto the point of the failure of the program, in case of program ending abruptly. So you will NOT need to redo the appts done uptil point of failure.\n\n\nLimitations of this program:\n\n1. Pyautogui doesn't work with dual monitor setup,i.e, it doesn't recognize the extended monitor. So please run only on 1920x1080 screens.\n2. If the client domain is slow, the program won't work properly.\n3. If an appt name such as "MRI Ankle" exists in the DCW but in the domain only one appt name exists starting with the same words, for e.g., "MRI Ankle Left", the program will add scheduling comments to the "MRI Ankle Left" appt and we won't be notified about the same. In result sheet it will say that it built it for "MRI Ankle" itself. Hence please double-check the appt mnemonics before entering in the excel sheet.\n4. This program does not associate scheduling comments in certain ambiguous cases such as multiple appointments starting with given appt mnemonic. These will need to be performed manually and this will be notified in the excel 'Results' sheet against the appt name.\n\n""")
    t.pack(side=TOP, fill=BOTH, expand=TRUE)
    t['font']=myFont4
    v.config(command=t.yview)




#define variable to read data from excel sheet
def sched_program():

    try :
        usr = os.environ['USERPROFILE']
        desk_excel = os.path.join((os.environ['USERPROFILE']),'OneDrive - Cerner Corporation','Desktop','Sched_comments.xlsx')

        df = pd.read_excel(desk_excel,sheet_name = "Sheet1")

        with pd.ExcelWriter(desk_excel,engine= "openpyxl", mode='a') as writer:
            writer.save()

        appt_list = df['Appointment Mnemonics'].values.tolist()

        #declare list variable to hold the results of each appt name
        output_list = []
        time.sleep(1)

        #prompt to get sched comment mnemonic
        scvalue = ''
        scvalue = pyautogui.prompt(text='Please enter the scheduling comment mnemonic to be moved:', title='Prompt' , default='')
        while scvalue == '':
            pyautogui.alert("Please enter a value to continue")
            scvalue = pyautogui.prompt(text='Please enter the scheduling comment mnemonic to be moved :', title='Prompt' , default='')
        if scvalue == None :
            sys.exit()
        else :
            pyautogui.alert(" you entered the value   " + str(scvalue) + "  !")

        x, y = pyautogui.locateCenterOnScreen('apptmnem.png',confidence=0.8)
        pyautogui.click(x, y)
        time.sleep(1)

        if pyautogui.locateOnScreen('apptname.png',confidence=0.8)!= None:
            x, y = pyautogui.locateCenterOnScreen('apptname.png',confidence=0.8)
            pyautogui.click(x, y)
            time.sleep(1)

            for name in range(len(appt_list)):  #typing appt name into the menmonic field in appt type tool to search for it
                pyautogui.write(appt_list[name])
                time.sleep(1)
                keyboard.press('enter')
                time.sleep(1)
                #if appt name is not found do the following:
                if pyautogui.locateOnScreen('apptnotfound.png',confidence=0.8)!=None:
                    time.sleep(1)
                    c, d = pyautogui.locateCenterOnScreen('apptnotfound.png',confidence=0.8)
                    pyautogui.click(c, d)
                    keyboard.press('enter')
                    time.sleep(1)
                    #write next to the name of current appt in excel that appt not apptnotfound
                    output_list.append('Appt mnemonic not found')

                elif pyautogui.locateOnScreen('synonyms.png',confidence=0.8)!=None: #finding multiple appts with similar names
                    time.sleep(1)
                    c, d = pyautogui.locateCenterOnScreen('cancel.png',confidence=0.8)
                    pyautogui.click(c, d)
                    time.sleep(1)
                    output_list.append('Not built since multiple similar appt names exist, please build manually')

                    #if appt name is found do this:
                else:
                    time.sleep(1)
                    if pyautogui.locateOnScreen('Comments1.png',confidence=0.8)!=None:
                        x, y = pyautogui.locateCenterOnScreen('Comments1.png',confidence=0.8)
                        pyautogui.doubleClick(x, y)
                        time.sleep(1)
                        #make the control go to 'scheduling comments' button
                        keyboard.press('down')
                        time.sleep(0.5)
                        keyboard.press('right')
                        time.sleep(0.5)
                        i=0
                        while i<5 :
                            keyboard.press('down')
                            time.sleep(0.5)
                            i=i+1
                        #if schedlunig comment already moved for all locations print 'already built' in excel sheet
                        #and click the 'clear' button and move to finding next appt name in excel sheet
                        time.sleep(1)
                        onscreen = pyautogui.locateOnScreen('selectedcomm.png',grayscale =False,confidence=0.8)
                        if onscreen !=None:
                            newx= onscreen[0] + onscreen[2]
                            newy= onscreen[1] + onscreen[3]
                            time.sleep(0.5)
                            pyautogui.moveTo(newx,newy)
                            pyautogui.click()

                            typewrite(str(scvalue),interval=.03)
                            time.sleep(2)
                            #if pyautogui.locateOnScreen('blue.png',confidence=0.9,grayscale =False) !=None:
                            if pyautogui.locateOnScreen('alllocs.png',confidence=0.8) !=None:
                                #comment not already moved so u can move it

                                onscreen = pyautogui.locateOnScreen('allcomms.png',confidence=0.8,grayscale =False)
                                if onscreen !=None:
                                    newx= onscreen[0] + onscreen[2]
                                    newy= onscreen[1] + onscreen[3]
                                    pyautogui.moveTo(newx,newy)
                                    time.sleep(1)
                                    pyautogui.click()
                                    time.sleep(1)
                                    typewrite(str(scvalue),interval=.03)
                                    time.sleep(1)
                                    if pyautogui.locateOnScreen('blue.png',confidence=0.9,grayscale =False) !=None:
                                        #move the comment and save
                                        x, y = pyautogui.locateCenterOnScreen('blue.png',confidence=0.9,grayscale =False)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        keyboard.press('tab')
                                        time.sleep(1)
                                        keyboard.press('enter')
                                        time.sleep(1)
                                        #loc = pyautogui.locateOnScreen('savebutton.png')
                                        #a = pyautogui.center(loc)
                                        #pyautogui.click(a.x, a.y)
                                        x, y = pyautogui.locateCenterOnScreen('savebutton.png',confidence=0.9)
                                        pyautogui.click(x, y)
                                        output_list.append('Success')
                                    else:

                                        #throw alert that comment is not found in domain and ask them to verify before running automation and exit program
                                        pyautogui.alert("the comment u provided is not in domain. plz verify and then re-run automation")
                                        sys.exit()
                            elif pyautogui.locateOnScreen('blue1.png',confidence=0.9,grayscale =False) !=None:
                                #comment is already moved clcik clear button

                                time.sleep(1)
                                if pyautogui.locateOnScreen('clearbutton.png',confidence=0.8) !=None:
                                    x, y = pyautogui.locateCenterOnScreen('clearbutton.png',confidence=0.8)
                                    pyautogui.click(x, y)
                                    output_list.append('Already exists in domain')
                                    time.sleep(1)
                                elif pyautogui.locateOnScreen('clear2.png',confidence=0.8) !=None:
                                    x, y = pyautogui.locateCenterOnScreen('clear2.png',confidence=0.8)
                                    pyautogui.click(x, y)
                                    output_list.append('Already exists in domain')
                                    time.sleep(1)
                                else:
                                    pyautogui.alert("clear button not found")
                                    sys.exit()
                            else:
                                time.sleep(1)

                                #comment can be selected and moved
                                onscreen = pyautogui.locateOnScreen('allcomms.png',confidence=0.8,grayscale =False)
                                if onscreen !=None:
                                    newx= onscreen[0] + onscreen[2]
                                    newy= onscreen[1] + onscreen[3]
                                    pyautogui.moveTo(newx,newy)
                                    pyautogui.click()
                                    time.sleep(1)
                                    typewrite(str(scvalue),interval=.03)
                                    time.sleep(1)
                                    onscreen1=pyautogui.locateOnScreen('blue.png',confidence=0.9,grayscale =False)
                                    if onscreen1 !=None:
                                        #move the comment and save
                                        #x, y = pyautogui.center(onscreen1)
                                        #y += 10
                                        time.sleep(1)
                                        x, y = pyautogui.locateCenterOnScreen('blue.png',confidence=0.9,grayscale =False)
                                        pyautogui.click(x, y)
                                        time.sleep(1)
                                        keyboard.press('tab')
                                        time.sleep(1)
                                        keyboard.press('enter')
                                        time.sleep(1)
                                        loc = pyautogui.locateOnScreen('savebutton.png',confidence=0.9)
                                        a = pyautogui.center(loc)
                                        pyautogui.click(a.x, a.y)
                                        output_list.append('Success')
                                        time.sleep(1)
                                    else:
                                        #throw alert that comment is not found in domain and ask them to verify before running automation and exit program
                                        pyautogui.alert("the comment u provided is not in domain. plz verify and then re-run automation")
                                        sys.exit()
                                else:
                                    pyautogui.alert("couldnt find all comments field")
                                    sys.exit()
                        else:
                            pyautogui.alert("couldnt find selected comments field")
                            sys.exit()
                    else:
                        pyautogui.alert("'Appt comments' button not visible")
                        sys.exit()
                df1 = pd.DataFrame(list(zip(appt_list,output_list)),columns=['Appt Mnemonic','Status'])

                with pd.ExcelWriter(desk_excel,engine= "openpyxl", mode='w') as writer:
                #with pd.ExcelWriter(desk_excel,engine= "openpyxl", mode='a') as writer:
                    df.to_excel(writer, index=False,sheet_name='Sheet1')
                    df1.to_excel(writer, index=False,sheet_name='Result')
                    writer.save()

            pyautogui.alert("Automation Complete!")

        else:
            pyautogui.alert("'Appointment Mnemonic' field not found on screen")
            sys.exit()
    except PermissionError:
        pyautogui.alert("Please close the excel workbook and then run automation.")
        sys.exit()

def stop1():
    # messagebox.showinfo("Pause","Program is paused")
    # pyautogui.alert("exiting program")
    sys.exit()

def buttonThread():
    t1=threading.Thread(target=sched_program)
    t1.daemon= True
    t1.start()

myFont = font.Font(family ="Rockwell",size=12)
myFont3 = font.Font(family ="Rockwell",size=14)
myFont2 = font.Font(family ="Rockwell",size=12, weight="bold")
myFont4 = font.Font(size=12)

label1 = Label(root, text="Welcome to Scheduling Comments Automation!")
button1 = Button(root, text="Instructions",command=my_instruct)
button2 = Button(root, text="Download Excel Template",command=create_excel)
button3 = Button(root, text="Run Automation",command=buttonThread)
button4 = Button(root, text="Quit",command=stop1)
label4 = Label(root,text="If you need to stop the running program due to any issues,\nplease open this window and click 'Quit' button")
# IMP!! If you need to stop the program during its exceution, please open this window and click 'Quit' button.
# The automation will stop immediately and all iterations upto the current will be saved on the excel sheet under 'Results' tab.
 # However the program will not restart from where you left off.
 # You will need to update the excel sheet and remove the names of appts whose execution was completed.

label1['font']=myFont3
button1['font']=myFont
button2['font']=myFont
button3['font']=myFont
button4['font']=myFont
label4['font']=myFont2

label1.pack(side= TOP, expand =True, padx=40, pady =50)
label4.pack(side= BOTTOM, padx=40, pady =50)
button1.pack(side=LEFT, padx=40, pady =50)
button2.pack(side=LEFT,padx=40,pady=50)
button3.pack(side=LEFT,padx=40,pady=50)
button4.pack(side=RIGHT,padx=40,pady=50)


root.mainloop()
