from pyautogui import *
import pyautogui
import time
import keyboard
import pandas as pd

#define variable to read data from excel sheet

try :
    #df = pd.read_excel(r'C:\Users\SS078074\OneDrive - Cerner Corporation\Desktop\Sched_comments.xlsx',sheet_name = "Sheet1")
    df = pd.read_excel(r'C:\Users\SS078074\OneDrive - Cerner Corporation\Desktop\User_Defined_Fields.xlsx',sheet_name="Sheet1")

    #with pd.ExcelWriter(r'C:\Users\SS078074\OneDrive - Cerner Corporation\Desktop\Sched_comments.xlsx',engine= "openpyxl", mode='a') as writer:
        #writer.save()

    #appt_list = df['Appointment Mnemonics'].values.tolist()
    f1 = df['Field Name/Prompt Description'].values.tolist()
    f2 = df['CDF/Unique Key'].values.tolist()
    f3 = df['PROMPT_TYPE'].values.tolist()

    #declare list variable to hold the results of each appt name
    output_list = []
    time.sleep(1)

    #prompt to get sched comment mnemonic
    scvalue = ''
    scvalue = pyautogui.prompt(text='What kind of User defined field you want to do add? Text?', title='Prompt' , default='')
    while scvalue == '':
        pyautogui.alert("Please enter a value to continue")
        scvalue = pyautogui.prompt(text='Please enter the scheduling comment mnemonic to be moved :', title='Prompt' , default='')
    if scvalue == None :
        exit()
    else :
        pyautogui.alert(" you entered the value   " + str(scvalue) + "  !")

    x, y = pyautogui.locateCenterOnScreen('Add.png')
    #pyautogui.click(x, y)
    time.sleep(1)

    if pyautogui.locateOnScreen('Add.png')!= None:
        x, y = pyautogui.locateCenterOnScreen('Add1.png')
        pyautogui.click(x, y)
        time.sleep(1)

        for name in range(len(f1)):  #typing appt name into the menmonic field in appt type tool to search for it
            pyautogui.write(f1[name])
            time.sleep(1)
            keyboard.press('tab')
            time.sleep(1)

            pyautogui.write(f2[name])
            time.sleep(1)
            keyboard.press('tab')

            time.sleep(1)
            pyautogui.write(f3[name])
            time.sleep(1)

            keyboard.press('tab')
            keyboard.press('enter')

            x, y = pyautogui.locateCenterOnScreen('Add1.png')
            pyautogui.click(x, y)
            time.sleep(1)

        pyautogui.alert("Automation Complete!")

    else:
        pyautogui.alert("Could not Add any")
        exit()

            #if appt name is not found do the following:
           if pyautogui.locateOnScreen('apptnotfound.png')!=None:
                time.sleep(1)
                c, d = pyautogui.locateCenterOnScreen('apptnotfound.png')
                pyautogui.click(c, d)
                keyboard.press('enter')
                time.sleep(1)
                #write next to the name of current appt in excel that appt not apptnotfound
                output_list.append('Appt mnemonic not found')

            elif pyautogui.locateOnScreen('synonyms.png')!=None: #finding multiple appts with similar names
                time.sleep(1)
                c, d = pyautogui.locateCenterOnScreen('cancel.png')
                pyautogui.click(c, d)
                time.sleep(1)
                output_list.append('Not built since multiple similar appt names exist, please build manually')

                #if appt name is found do this:
            else:
                time.sleep(1)
                if pyautogui.locateOnScreen('Comments1.png',confidence=0.8)!=None:
                    x, y = pyautogui.locateCenterOnScreen('Comments1.png')
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
                        time.sleep(0.25)
                        i=i+1
                    #if schedlunig comment already moved for all locations print 'already built' in excel sheet
                    #and click the 'clear' button and move to finding next appt name in excel sheet
                    time.sleep(1)
                    onscreen = pyautogui.locateOnScreen('selectedcomm.png',grayscale =False)
                    if onscreen !=None:
                        newx= onscreen[0] + onscreen[2]
                        newy= onscreen[1] + onscreen[3]
                        time.sleep(0.5)
                        pyautogui.moveTo(newx,newy)
                        pyautogui.click()

                        typewrite(str(scvalue),interval=.03)
                        time.sleep(2)
                        #if pyautogui.locateOnScreen('blue.png',confidence=0.9,grayscale =False) !=None:
                        if pyautogui.locateOnScreen('alllocs.png') !=None:
                            #comment not already moved so u can move it
                            print("if part1")
                            onscreen = pyautogui.locateOnScreen('allcomms.png',grayscale =False)
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
                                    time.sleep(0.25)
                                    keyboard.press('tab')
                                    time.sleep(0.25)
                                    keyboard.press('enter')
                                    time.sleep(1)
                                    #loc = pyautogui.locateOnScreen('savebutton.png')
                                    #a = pyautogui.center(loc)
                                    #pyautogui.click(a.x, a.y)
                                    x, y = pyautogui.locateCenterOnScreen('savebutton.png')
                                    pyautogui.click(x, y)
                                    output_list.append('Success')
                                else:

                                    #throw alert that comment is not found in domain and ask them to verify before running automation and exit program
                                    pyautogui.alert("the comment u provided is not in domain. plz verify and then re-run automation")
                                    exit()
                        elif pyautogui.locateOnScreen('blue1.png',confidence=0.9,grayscale =False) !=None:
                            #comment is already moved clcik clear button
                            print("if part 2")
                            time.sleep(1)
                            if pyautogui.locateOnScreen('clearbutton.png') !=None:
                                x, y = pyautogui.locateCenterOnScreen('clearbutton.png')
                                pyautogui.click(x, y)
                                output_list.append('Already exists in domain')
                                time.sleep(1)
                            elif pyautogui.locateOnScreen('clear2.png') !=None:
                                x, y = pyautogui.locateCenterOnScreen('clear2.png')
                                pyautogui.click(x, y)
                                output_list.append('Already exists in domain')
                                time.sleep(1)
                            else:
                                pyautogui.alert("clear button not found")
                                exit()
                        else:
                            time.sleep(1)
                            print("if part 3")
                            #comment can be selected and moved
                            onscreen = pyautogui.locateOnScreen('allcomms.png',grayscale =False)
                            if onscreen !=None:
                                newx= onscreen[0] + onscreen[2]
                                newy= onscreen[1] + onscreen[3]
                                pyautogui.moveTo(newx,newy)
                                pyautogui.click()
                                time.sleep(0.5)
                                typewrite(str(scvalue),interval=.03)
                                time.sleep(0.5)
                                onscreen1=pyautogui.locateOnScreen('blue.png',confidence=0.9,grayscale =False)
                                if onscreen1 !=None:
                                    #move the comment and save
                                    #x, y = pyautogui.center(onscreen1)
                                    #y += 10
                                    time.sleep(1)
                                    x, y = pyautogui.locateCenterOnScreen('blue.png',confidence=0.9,grayscale =False)
                                    pyautogui.click(x, y)
                                    time.sleep(0.5)
                                    keyboard.press('tab')
                                    time.sleep(0.5)
                                    keyboard.press('enter')
                                    time.sleep(0.5)
                                    loc = pyautogui.locateOnScreen('savebutton.png')
                                    a = pyautogui.center(loc)
                                    pyautogui.click(a.x, a.y)
                                    output_list.append('Success')
                                else:
                                    #throw alert that comment is not found in domain and ask them to verify before running automation and exit program
                                    pyautogui.alert("the comment u provided is not in domain 2. plz verify and then re-run automation")
                                    exit()
                            else:
                                pyautogui.alert("couldnt find all comments field")
                                exit()
                    else:
                        pyautogui.alert("couldnt find selected comments field")
                        exit()
                else:
                    pyautogui.alert("'Appt comments' button not visible")
                    exit()
            df1 = pd.DataFrame(list(zip(appt_list,output_list)),columns=['Appt Mnemonic','Status'])

except PermissionError:
    pyautogui.alert("Please close the excel workbook and then run automation.")
    exit()
