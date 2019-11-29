# -*- coding: cp936 -*-
import os
import sys
import subprocess
import time
from xlrd import open_workbook
import xlrd
import xlsxwriter
import xlwt
from xlwt import Workbook
import glob
import serial
import os

sdir=os.getcwd()
zadir=str(glob.glob(sdir + "/" + "report.xlsx"))
cutt=zadir.split("/")
beforee=cutt[-1]
finals=beforee.replace("']", "")
global m
m=4

if finals==str("report.xlsx"):
    os.remove("report.xlsx")

else:
    print("")


def sendCommand(com):
    ser=serial.Serial('COM'+COM, baudrate=9600, timeout=.1, rtscts=0)
    ser.write(com+"\r\n")
    time.sleep(2)
    ret = []
    while ser.inWaiting() > 0:
            msg = ser.readline().strip()
            msg = msg.replace("\r","")
            msg = msg.replace("\n","")
            if msg!="":
                ret.append(msg)
			
    return ret
    ser.close()
    
#######Fuction to register on to the 4G network###################
    
def check():
    print("check")
    os.system("adb shell cmd statusbar expand-notifications")

    from com.dtmilano.android.viewclient import ViewClient

    device, serialno=ViewClient.connectToDeviceOrExit()
    vc=ViewClient(device=device, serialno=serialno)

    Icon=vc.findViewsWithAttribute("resource-id","com.android.systemui:id/mobile_combo_card")
    time.sleep(1)
    vc.dump()
    Mode=(Icon[0].getContentDescription())
    List=list(Mode)
    L1=List[0]
    L2=List[1]
    global String
    String= L1+ L2
    print(String)
    Split=Mode.split(' ')
    
    L3=Split[1]
    
    L4=Split[2]
    combine=L3+" " + L4
    print(combine)
          
    sets=String+" " + combine
    print(sets)
    
    global gets
    gets=String+" "+ L3
    print(gets)



def refresh():
    try:
        cdir=os.getcwd()
        adir=str(glob.glob(cdir + "/" + "logcat.txt"))
        cut=adir.split("/")
        before=cut[-1]
        global final
        final=before.replace("']", "")
        
        if final=="logcat.txt":
            print("Log file found Removing now.....")
            os.remove(final)
        else:
            print("No log File found....")
            
        
        from com.dtmilano.android.viewclient import ViewClient
        device, serialno=ViewClient.connectToDeviceOrExit()
        vc=ViewClient(device=device, serialno=serialno)
        device.shell("input keyevent KEYCODE_HOME")
        vc.dump()
        for i in range(6):      
            from com.dtmilano.android.viewclient import ViewClient
            device, serialno=ViewClient.connectToDeviceOrExit()
            vc=ViewClient(device=device, serialno=serialno)

            if vc.findViewWithText("Settings"):
                vc.dump()
                vc.findViewWithText("Settings").touch()
                vc.dump()
                vc.findViewWithText("Wireless & networks").touch()
                vc.dump()
                vc.findViewById("android:id/switch_widget").touch()
                vc.dump()
                time.sleep(3)
                vc.findViewById("android:id/switch_widget").touch()
                vc.dump()
                device.shell("input keyevent KEYCODE_HOME")
                vc.dump()
                break
            else:
                os.system("adb shell input swipe 876 856 102 949 ")
                time.sleep(1)
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithText("Settings"):
                    vc.dump()
                    vc.findViewWithText("Settings").touch()
                    vc.dump()
                    vc.findViewWithText("Wireless & networks").touch()
                    vc.dump()
                    vc.findViewById("android:id/switch_widget").touch()
                    vc.dump()
                    time.sleep(3)
                    vc.findViewById("android:id/switch_widget").touch()
                    vc.dump()
                    device.shell("input keyevent KEYCODE_HOME")
                    vc.dump()
                    break
                
      
        
        os.system("adb logcat -d time >> logcat.txt")
        
        time.sleep(4)
        file=open("logcat.txt", "r")
        lines= file.read()
        if "ims registered= true" in lines:
            sheet1.write(j,0,PLMN)
            sheet1.write(j,1,"PASS")
            time.sleep(1)
            os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"Registration.png")
            time.sleep(1)
            os .system("adb pull /sdcard/Automation/"+PLMN+"Registration.png "+fin+PLMN+"Registration.png" )
            time.sleep(1)
            print("writing1")
            file.close()
        else:
            sheet1.write(j,0,PLMN)
            sheet1.write(j,1,"FAIL")
            time.sleep(1)
            os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"Registration.png")
            time.sleep(1)
            os .system("adb pull /sdcard/Automation/"+PLMN+"Registration.png "+fin+PLMN+"Registration.png" )
            time.sleep(1)
            print("writing1")
            file.close()

    except:
        print("An Exception occured......")
        
def FailRegister():
    try:
        from com.dtmilano.android.viewclient import ViewClient
        device, serialno=ViewClient.connectToDeviceOrExit()
        vc=ViewClient(device=device, serialno=serialno)
        vc.findViewById("android:id/button1").touch()
        vc.dump()
        os.system("adb shell input keyevent KEYCODE_HOME")
        time.sleep(2)
        os.system("adb shell pm clear com.android.phone")
        time.sleep(1)
        print(sendCommand("AT+CFUN=0"))
        time.sleep(3)
        print(sendCommand("AT+CFUN=1"))
        time.sleep(3)
        print(sendCommand("AT+CFUN=1,1"))
        time.sleep(45)
        from com.dtmilano.android.viewclient import ViewClient
        device, serialno=ViewClient.connectToDeviceOrExit()
        vc=ViewClient(device=device, serialno=serialno)
        if vc.findViewWithText("MAI TEST"):
            vc.dump()
            if vc.findViewById("android:id/button1"): 
                vc.dump()
                time.sleep(1)
                vc.findViewById("android:id/button1").touch()
                vc.dump()
                print("clicked")
                check()
                if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                    os.system("adb shell cmd statusbar collapse")
                    refresh()
                
            else:
                print("")

        else:
            for i in range(4):
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithText("Settings"):
                    vc.dump()
                    vc.findViewWithText("Settings").touch()
                    vc.dump()
                    vc.findViewWithText("Wireless & networks").touch()
                    vc.dump()
                    vc.findViewWithText("Mobile network").touch()
                    vc.dump()
                    break
                else:
                    os.system("adb shell input swipe 876 856 102 949 ")
                    time.sleep(1)
                    from com.dtmilano.android.viewclient import ViewClient
                    device, serialno=ViewClient.connectToDeviceOrExit()
                    vc=ViewClient(device=device, serialno=serialno)
                    if vc.findViewWithText("Settings"):
                        vc.dump()
                        vc.findViewWithText("Settings").touch()
                        vc.dump()
                        vc.findViewWithText("Wireless & networks").touch()
                        vc.dump()
                        vc.findViewWithText("Mobile network").touch()
                        vc.dump()
                        break        
            from com.dtmilano.android.viewclient import ViewClient
            device, serialno=ViewClient.connectToDeviceOrExit()
            vc=ViewClient(device=device, serialno=serialno)
            vc.findViewWithText("Carrier").touch()
            time.sleep(2)
            vc.dump()
            vc.findViewById("android:id/switch_widget").touch()
            vc.dump()
            vc.findViewById("android:id/button1").touch()
            time.sleep(50)
            vc.dump()
            check() 
            if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                
                os.system("adb shell cmd statusbar collapse")
                refresh()

            else:
                os.system("adb shell cmd statusbar collapse")
                time.sleep(1)
                if vc.findViewWithText(u"中国联通 4G"):
                    vc.dump()
                    vc.findViewWithText(u"中国联通 4G").touch()
                    time.sleep(2)
                    vc.dump()
                
                    if vc.findViewWithText("Network registration failed"):
                        vc.dump()
                        vc.findViewById("android:id/button1").touch()
                        vc.dump()
                        sheet1.write(j,0,PLMN)
                        sheet1.write(j,1,"FAIL")
                        time.sleep(1)
                        os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"Registration.png")
                        time.sleep(1)
                        os .system("adb pull /sdcard/Automation/"+PLMN+"Registration.png "+fin+PLMN+"Registration.png" )
                        time.sleep(1)

                    elif vc.findViewWithText("MAI TEST"):
                        vc.dump()
                        vc.findViewById("android:id/button1").touch()
                        vc.dump()
                        refresh()
                

                elif vc.findViewWithText(u"中国联通 LTE"):
                    vc.dump()
                    vc.findViewWithText(u"中国联通 LTE").touch()
                    time.sleep(2)
                    vc.dump()
                    if vc.findViewWithText("Network registration failed"):
                        vc.dump()
                        vc.findViewById("android:id/button1").touch()
                        vc.dump()
                        sheet1.write(j,0,PLMN)
                        sheet1.write(j,1,"FAIL")
                        time.sleep(1)
                        os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"Registration.png")
                        time.sleep(1)
                        os .system("adb pull /sdcard/Automation/"+PLMN+"Registration.png "+fin+PLMN+"Registration.png" )
                        time.sleep(1)

                    elif vc.findViewWithText("MAI TEST"):
                        vc.dump()
                        vc.findViewById("android:id/button1").touch()
                        vc.dump()
                        refresh()
                
            
    except:
        print("Found an exception......")

def PassRegister():
    try:
        vc.findViewById("android:id/button1").touch()
        vc.dump()
        refresh()
    except:
        print("")

    
def Network():
    try:
        from com.dtmilano.android.viewclient import ViewClient
        device, serialno=ViewClient.connectToDeviceOrExit()
        vc=ViewClient(device=device, serialno=serialno)
        if vc.findViewWithText("MAI TEST"):
            vc.dump()
            if vc.findViewById("android:id/button1"):
                
                vc.dump()
                time.sleep(1)
                vc.findViewById("android:id/button1").touch()
                vc.dump()
                print("clicked")
            else:
                time.sleep(15)
                
        else:
            print""
        check()
        
        if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
            print("4g")
            os.system("adb shell cmd statusbar collapse")
            refresh()
            
            

        elif String==str("3G") or combine==str("No signal") or sets==str("4G No signal"):
            print("3g")
            time.sleep(1)
            cdir=os.getcwd()
            adir=str(glob.glob(cdir + "/" + "logcat.txt"))
            cut=adir.split("/")
            before=cut[-1]
            final=before.replace("']", "")
            if final=="logcat.txt":
                print("Log file found Removing now.....")
                os.remove(final)

            else:
                print("No log file found...........")
                
            time.sleep(1)
                
            os.system("adb shell cmd statusbar collapse")
            
            
            for i in range(4):
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithText("Settings"):
                    vc.dump()
                    vc.findViewWithText("Settings").touch()
                    vc.dump()
                    vc.findViewWithText("Wireless & networks").touch()
                    vc.dump()
                    vc.findViewWithText("Mobile network").touch()
                    vc.dump()
                    break
                else:
                    os.system("adb shell input swipe 876 856 102 949 ")
                    time.sleep(1)
                    from com.dtmilano.android.viewclient import ViewClient
                    device, serialno=ViewClient.connectToDeviceOrExit()
                    vc=ViewClient(device=device, serialno=serialno)
                    if vc.findViewWithText("Settings"):
                        vc.dump()
                        vc.findViewWithText("Settings").touch()
                        vc.dump()
                        vc.findViewWithText("Wireless & networks").touch()
                        vc.dump()
                        vc.findViewWithText("Mobile network").touch()
                        vc.dump()
                        break        

            from com.dtmilano.android.viewclient import ViewClient
            device, serialno=ViewClient.connectToDeviceOrExit()
            vc=ViewClient(device=device, serialno=serialno)
            vc.findViewWithText("Carrier").touch()        
            vc.dump()
            
            try:
                if not vc.findViewWithText("Automatic") and not vc.findViewWithText("Disable auto-select"):
                    time.sleep(50)
                    vc.dump()
                    check()
                    os.system("adb shell cmd statusbar collapse")
                    if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                        refresh()
                else:
                    time.sleep(1)
                    if vc.findViewWithText(u"中国联通 4G"):
                        vc.dump()
                        vc.findViewWithText(u"中国联通 4G").touch()
                        time.sleep(8)
                        vc.dump()   
                        
                        if not vc.findViewWithText("MAI TEST"):
                            time.sleep(32)
                            vc.dump()
                            
                            
                        if vc.findViewWithText("MAI TEST"):
                            vc.dump()
                            PassRegister()
                            
                        elif vc.findViewWithText("Network registration failed"):
                            vc.dump()
                            FailRegister()
                            
                    elif vc.findViewWithText(u"中国联通 LTE"):
                        vc.dump()
                        vc.findViewWithText(u"中国联通 LTE").touch()
                        time.sleep(8)
                        vc.dump()

                        if not vc.findViewWithText("MAI TEST"):
                            time.sleep(32)
                            vc.dump()
                            

                        if vc.findViewWithText("MAI TEST"):
                            vc.dump()
                            PassRegister()

                        elif vc.findViewWithText("Network registration failed"):
                            vc.dump()
                            FailRegister()

                    else:
                        sheet1.write(j,0,PLMN)
                        sheet1.write(j,1,"FAIL")
                        
                    check()
                    os.system("adb shell cmd statusbar collapse")
                    if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                        refresh()

                    else:
                        print("")
                
            except:
                print("An Exception occured......")

                
                    ##############################################
                        
           

                    #######################################################
               
            try:

                if vc.findViewWithText("Automatic"):
                    time.sleep(2)
                    vc.dump()
                    
                    vc.findViewById("android:id/switch_widget").touch()
                    vc.dump()
                    vc.findViewById("android:id/button1").touch()
                    time.sleep(50)
                    vc.dump()
                    
                    check()
                    os.system("adb shell cmd statusbar collapse")
                    if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                        refresh()
                        
                    else:
                        
                        time.sleep(1)
                        if vc.findViewWithText(u"中国联通 4G"):
                            vc.dump()
                            vc.findViewWithText(u"中国联通 4G").touch()
                            time.sleep(8)
                            
                            if not vc.findViewWithText("MAI TEST") :
                                time.sleep(32)
                                vc.dump()
                                
                                
                            if vc.findViewWithText("MAI TEST"):
                                vc.dump()
                                PassRegister()
                                
                            elif vc.findViewWithText("Network registration failed"):
                                vc.dump()
                                FailRegister()
                                
                            
                            
                        elif vc.findViewWithText(u"中国联通 LTE"):
                            vc.dump()
                            
                           
                            vc.findViewWithText(u"中国联通 LTE").touch()
                            time.sleep(8)
                            vc.dump()
                            

                            if not vc.findViewWithText("MAI TEST"):
                                vc.dump()
                                time.sleep(32)

                            if vc.findViewWithText("MAI TEST"):
                                vc.dump()
                                PassRegister()

                            
                            elif vc.findViewWithText("Network registration failed"):
                                vc.dump()
                                FailRegister()

                        else:
                            sheet1.write(j,0,PLMN)
                            sheet1.write(j,1,"FAIL")
                        check()
                        os.system("adb shell cmd statusbar collapse")
                        if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                            refresh()

                        else:
                            print("")

            except:
                
                print("An Exception occured......")


            try:
                if vc.findViewWithText("Disable auto-select"):
                    time.sleep(1)
                    vc.dump()
                    
                    
                    vc.findViewById("android:id/button1").touch()
                    time.sleep(50)
                    vc.dump()
                    check()
                    os.system("adb shell cmd statusbar collapse")
                    if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                        refresh()

                    else:
                        
                        time.sleep(1)
                        if vc.findViewWithText(u"中国联通 4G"):
                            vc.dump()
                            vc.findViewWithText(u"中国联通 4G").touch()
                            time.sleep(8)
                            vc.dump()
                            
                            if not vc.findViewWithText("MAI TEST") :
                                time.sleep(32)
                                vc.dump()
                                
                                
                            if vc.findViewWithText("MAI TEST"):
                                vc.dump()
                                PassRegister()
                                
                            elif vc.findViewWithText("Network registration failed"):
                                vc.dump()
                                FailRegister()
                


                        elif vc.findViewWithText(u"中国联通 LTE"):
                            vc.dump()
                            vc.findViewWithText(u"中国联通 LTE").touch()
                            
                            time.sleep(40)
                            vc.dump()
                           

                            if not vc.findViewWithText("MAI TEST"):
                                time.sleep(32)
                                vc.dump()
                                

                            
                            if vc.findViewWithText("MAI TEST"):
                                vc.dump()
                                PassRegister()

                            
                            elif vc.findViewWithText("Network registration failed"):
                                vc.dump()
                                FailRegister()

                        else:
                            sheet1.write(j,0,PLMN)
                            sheet1.write(j,1,"FAIL")

                        check()
                        os.system("adb shell cmd statusbar collapse")
                        if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                            refresh()
                        else:
                            print("")

            except:
                print("An Exception occured......")
            
    except:

        print("An Exception occured......")

###########################Call#################################################################################

def Call():
    
    try:
        cdir=os.getcwd()
        adir=str(glob.glob(cdir + "/" + "logcat.txt"))
        cut=adir.split("/")
        before=cut[-1]
        global final
        final=before.replace("']", "")
        if final=="logcat.txt":
            print("Log file found Removing now.....")
            os.remove(final)

        else:
            print("No log file found...........")
        os.system("adb shell pm clear com.android.phone")
        time.sleep(10)
        check()
        time.sleep(1)
        os.system("adb shell cmd statusbar collapse")
        time.sleep(1)
        print("bye")
        if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
            try:

                Call="adb shell am start -a android.intent.action.CALL -d tel:" + str(ISDN)
                output=subprocess.Popen(Call,stdout=subprocess.PIPE).communicate()[0]     

                time.sleep(15)
                os.system("adb logcat -d time >> logcat.txt")
                time.sleep(4)
                file=open("logcat.txt", "r")
                lines= file.read()
                

                if "CallState DIALING -> ACTIVE" in lines and "isVolteCall()" in lines:
                    time.sleep(1)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"call.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"call.png "+fin+PLMN+"call.png" )
                    time.sleep(6)
                    os.system("adb shell input keyevent KEYCODE_ENDCALL")
                
                    sheet1.write(j,2,"PASS")
                    
                    file.close()



                elif "CallState DIALING -> DISCONNECTED" in lines:
                    time.sleep(1)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"call.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"call.png "+fin+PLMN+"call.png" )
                    time.sleep(2)
                    sheet1.write(j,2,"FAIL")
                    #############Write it in excel
                   
                    file.close()

                elif "CallState DIALING -> DIALING" in lines:
                    time.sleep(3)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"call.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"call.png "+fin+PLMN+"call.png" )
                    time.sleep(2)
                    os.system("adb shell input keyevent KEYCODE_ENDCALL")
                    
                    sheet1.write(j,2,"FAIL")#############Write it in excel
                    
                    file.close()

                elif "CallState DIALING -> ACTIVE" in lines:
                    
                    time.sleep(5)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"call.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"call.png "+fin+PLMN+"call.png" )
                    time.sleep(2)
                    os.system("adb shell input keyevent KEYCODE_ENDCALL")
                    sheet1.write(j,2,"FAIL")
                    file.close()

                elif "CallState DIALING" not in lines:
                    time.sleep(3)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"call.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"call.png "+fin+PLMN+"call.png" )
                    time.sleep(2)
                    os.system("adb shell input keyevent KEYCODE_ENDCALL")
                    sheet1.write(j,0,PLMN)
                    sheet1.write(j,2,"FAIL")#############Write it in excel
                    
                    file.close()
            
            except:
                print("Exception while dialing.......")
                

        else: 
            sheet1.write(j,2,"FAIL")
            
        os.system("adb shell input keyevent KEYCODE_WAKEUP")
        time.sleep(1)
        os.system("adb shell pm clear com.android.contacts")
        time.sleep(1)
        os.system("adb shell input keyevent KEYCODE_HOME")

            

        from com.dtmilano.android.viewclient import ViewClient
        device, serialno=ViewClient.connectToDeviceOrExit()
        vc=ViewClient(device=device, serialno=serialno)

        try:
            vc.findViewWithContentDescription("Phone").touch()
            vc.dump()
            time.sleep(1)
            os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"callog.png")
            time.sleep(2)
            os .system("adb pull /sdcard/Automation/"+PLMN+"callog.png"+fin+PLMN+"callog.png")
            time.sleep(2)
            ###############TakeScreenshot
            os.system("adb shell input keyevent KEYCODE_HOME")

        except:
            print("Exception occured while finding element")
       
    except:
        print("An Exception occured......")
    
###############################################SMS########################################################################
        
def SMS():
    try:   
        cdir=os.getcwd()
        adir=str(glob.glob(cdir + "/" + "logcat.txt"))
        cut=adir.split("/")
        before=cut[-1]
        global final
        final=before.replace("']", "")
        os.system("adb shell pm clear com.android.phone")
        time.sleep(10)
        check()
        os.system("adb shell cmd statusbar collapse")
        time.sleep(1)
        if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):

            if final=="logcat.txt":
                print("Log file found Removing now.....")
                os.remove(final)
                
            else:
                print("No log file found..........")
            
            try:
                    
                time.sleep(1)
                os.system("adb shell pm clear com.google.android.apps.messaging")
                
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithContentDescription("Messages"):
                    vc.findViewWithContentDescription("Messages").touch()
                    time.sleep(1)
                    vc.dump()
                else:
                    device.startActivity("com.huawei.android.launcher/com.huawei.android.launcher.unihome.UniHomeLauncher")
                    time.sleep(1)
                    vc.dump()
                    
                vc.findViewWithContentDescription("Start chat").touch()
                time.sleep(1)
                vc.dump()
                vc.findViewById("com.google.android.apps.messaging:id/recipient_text_view").touch()
                time.sleep(2)
                vc.dump()
                device.shell("input text " + str(ISDN))
                vc.dump()           
                time.sleep(3)
                vc.findViewById("com.google.android.apps.messaging:id/contact_picker_create_group").touch()
                time.sleep(1)
                vc.dump()
                vc.findViewById("com.google.android.apps.messaging:id/compose_message_text").setText("Hi")
                time.sleep(1)
                vc.dump()
                vc.findViewById("com.google.android.apps.messaging:id/send_message_button_container").touch()
                time.sleep(4)
                vc.dump()
                os.system("adb logcat -d time >> logcat.txt")
                time.sleep(6)
                file=open("logcat.txt", "r")
                lines= file.read()

                if "BugleDataMode: Processing changed messages for 357" in lines or "Done sending SMS message{id:357} conversation{id:50}, status: MANUAL_RETRY" in lines or "process from ProcessSentMessageAction due to sms_send failure with queues:" in lines:
                    
                    sheet1.write(j,3,"FAIL")
                    time.sleep(1)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"sms.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"sms.png "+fin+PLMN+"sms.png" )
                    time.sleep(1)
                    file.close()

                elif "status: SUCCEEDED" in lines:
                    
                    sheet1.write(j,3,"PASS")
                    time.sleep(1)
                    os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"sms.png")
                    time.sleep(1)
                    os .system("adb pull /sdcard/Automation/"+PLMN+"sms.png "+fin+PLMN+"sms.png" )
                    time.sleep(1)
                    file.close()

                    #Take screenshot
            except:
                
                print("Exception while sending SMS..............")
                
            time.sleep(1)
            os.system("adb shell pm clear com.google.android.apps.messaging")
            time.sleep(1)
            
            
        
        else:
            sheet1.write(j,3,"FAIL")

        os.system("adb shell input keyevent KEYCODE_HOME")
        time.sleep(2)
        for i in range(4):
            from com.dtmilano.android.viewclient import ViewClient
            device, serialno=ViewClient.connectToDeviceOrExit()
            vc=ViewClient(device=device, serialno=serialno)
            if vc.findViewWithText("Settings"):
                vc.dump()
                vc.findViewWithText("Settings").touch()
                vc.dump()
                vc.findViewWithText("Wireless & networks").touch()
                vc.dump()
                vc.findViewWithText("Mobile network").touch()
                vc.dump()
                break
            else:
                os.system("adb shell input swipe 876 856 102 949 ")
                time.sleep(1)
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithText("Settings"):
                    vc.dump()
                    vc.findViewWithText("Settings").touch()
                    vc.dump()
                    vc.findViewWithText("Wireless & networks").touch()
                    vc.dump()
                    vc.findViewWithText("Mobile network").touch()
                    vc.dump()
                    break        

        from com.dtmilano.android.viewclient import ViewClient
        device, serialno=ViewClient.connectToDeviceOrExit()
        vc=ViewClient(device=device, serialno=serialno)
        os.system("adb shell screencap -p /sdcard/Automation/"+PLMN+"Settings.png")
        time.sleep(1)
        os .system("adb pull /sdcard/Automation/"+PLMN+"Settings.png "+fin+PLMN+"Settings.png" )
        time.sleep(1)
    

    except:
        print("An Exception occured......")


###############Above is the set of fuctions#####################################

def reverse(str):
    strlist=[]
    for i in range(len(str)-1):
        if i%2==0:
            strlist.append(str[i+1])
        else:
            strlist.append(str[i-1])
    if len(str)%2==0:
        strlist.append(str[len(str)-2])
    else:
        strlist.append(str[len(str)-1])
    return ''.join(strlist)


wb=open_workbook('PLMN.xlsx')
sheet=wb.sheet_by_index(0)
rows=sheet.nrows
coloumn=sheet.ncols
wb1 = Workbook()
sheet1 = wb1.add_sheet('Sheet')
sheet1.write(0, 0, 'PLMN')
sheet1.write(0, 1, 'IMS Registration')
sheet1.write(0, 2, 'MO Call over IMS')
sheet1.write(0,3,'MO SMS')
cwd=os.getcwd()+'\\Screenshots\\'
getz=str(cwd.split())
va=getz.replace("['","")
fin=va.replace("']","")

from com.dtmilano.android.viewclient import ViewClient
device, serialno=ViewClient.connectToDeviceOrExit()
vc=ViewClient(device=device, serialno=serialno)

imsi=raw_input("Enter the IMSI that you want to make call and SMS.....")
COM=raw_input("Please enetr the PCUI port")
try:
    for j in range(0,rows):    
        PLMN=sheet.cell_value(j,0)
        first=PLMN    
        j=j+1

        if "755" in PLMN:
            mccmnc=PLMN[:5]
            print("5 digit imsi")
            newstrg=imsi[6:]
            
            inISDN=int("+91" + newstrg)
            ISDN="+" + str(inISDN)
            print(ISDN)
            print(sendCommand("AT^logport=0"))
            time.sleep(1)
            print(sendCommand("AT^Modem=1"))
            time.sleep(1)

            command3='com::send {at+csim=14,"00A40004026FAD";}\n'
            command4='com::send {at+csim=10,"00B000000B";}\n'
            command5='com::send {at+csim=18,"00D600000400000102";}\n'
            command6='com::send {at+csim=14,"00A40004026F07";}\n'
            command7='com::send {at+csim=10,"00B000000B";}\n'
            command8='com::send {at+csim=28,"00D600000908'
            command8+=reverse('9'+ PLMN)
            command8+='";}\n'
            commandList=[]
            commandList.append('com::send {at+csim=14,"00A40004023F00";}\n')
            commandList.append('com::recv OK\n')
            commandList.append('com::send {at+csim=14,"00A40004027FF0";}\n')
            commandList.append('com::recv OK\n')
            commandList.append('com::send {at+csim=26,"0020000A083030303030303030";}\n')
            commandList.append('com::recv OK\n')
            commandList.append(command3)
            commandList.append('com::recv OK\n')
                        #commandList.append(command2)
                        #commandList.append('com::recv OK\n')
            commandList.append(command5)
            commandList.append('com::recv OK\n')  
            commandList.append(command6)
            commandList.append('com::recv OK\n')
                        #commandList.append(command5)
                        #commandList.append('com::recv OK\n')
            commandList.append(command8)
            commandList.append('com::recv OK\n')
            file_object=open("script\csim.tcl","w")
            file_object.writelines(commandList)
            file_object.close()

            os.system("ATTv1.05.exe")
            time.sleep(10)
            fmod=open("script\csim.log","r")
            lineIn=fmod.readlines()
            for strIn in lineIn:
                strIn=strIn.strip("\n")
                if strIn.find("ERROR")>0:
                    print "\n"
                    print "executeATTScript failed."

                    fmod.close()
            fmod.close()
            filx=open("log.txt", "r")
            linez= filx.read()
    
            
            if str("fail: 1") in linez:
                sheet1.write(j,0,PLMN)
                sheet1.write(j,1,"FAIL to write IMSI")
                sheet1.write(j,2,"FAIL to write IMSI")
                sheet1.write(j,3,"FAIL to write IMSI")
                wb1.save('report.xls')

            elif str("pass") in linez:
                os.system("adb shell pm clear com.android.phone")
                os.system("adb reboot")
                time.sleep(85)
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithText("MAI TEST"):
                    vc.dump()
                    if vc.findViewById("android:id/button1"): 
                        vc.dump()
                        time.sleep(1)
                        vc.findViewById("android:id/button1").touch()
                        vc.dump()
                        print("clicked")
                        Network()        
                        Call()    
                        SMS()   
                        wb1.save('report.xls')
                    else:
                        print("")
                    

                else:
                    check()
                    os.system("adb shell cmd statusbar collapse")
                    if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                        refresh()
                        Call()    
                        SMS()   
                        wb1.save('report.xls')
                    else:
                        print(sendCommand("AT+CFUN=0"))
                        time.sleep(3)
                        print(sendCommand("AT+CFUN=1"))
                        time.sleep(3)
                        print(sendCommand("AT+CFUN=1,1"))
                        time.sleep(45)
                        Network()        
                        Call()    
                        SMS()   
                        wb1.save('report.xls')
            filx.close()


            

        elif "755" not in PLMN:
            mccmnc=imsi[:6]
            print("6 digit imsi")
            inewstrg=imsi[6:]
            ninISDN=int("+917" + inewstrg)
            ISDN="+" + str(ninISDN)
            print(ISDN)
            print(sendCommand("AT^logport=0"))
            time.sleep(1)
            print(sendCommand("AT^Modem=1"))
            time.sleep(1)
           
            command3='com::send {at+csim=14,"00A40004026FAD";}\n'
            command4='com::send {at+csim=10,"00B000000B";}\n'
            command5='com::send {at+csim=18,"00D600000400000103";}\n'
            command6='com::send {at+csim=14,"00A40004026F07";}\n'
            command7='com::send {at+csim=10,"00B000000B";}\n'
            command8='com::send {at+csim=28,"00D600000908'
            command8+=reverse('9'+PLMN)
            command8+='";}\n'
            commandList=[]
           # commandList.append(command1)
           # commandList.append('com::recv OK\n')
           # commandList.append(command2)
           # commandList.append('com::recv OK\n')
            commandList.append('com::send {at+csim=14,"00A40004023F00";}\n')
            commandList.append('com::recv OK\n')
            commandList.append('com::send {at+csim=14,"00A40004027FF0";}\n')
            commandList.append('com::recv OK\n')
            commandList.append('com::send {at+csim=26,"0020000A083030303030303030";}\n')
            commandList.append('com::recv OK\n')
            commandList.append(command3)
            commandList.append('com::recv OK\n')
                        #commandList.append(command2)
                        #commandList.append('com::recv OK\n')
            commandList.append(command5)
            commandList.append('com::recv OK\n')  
            commandList.append(command6)
            commandList.append('com::recv OK\n')
                        #commandList.append(command5)
                        #commandList.append('com::recv OK\n')
            commandList.append(command8)
            commandList.append('com::recv OK\n')
            file_object=open("script\csim.tcl","w")
            file_object.writelines(commandList)
            file_object.close()

            os.system("ATTv1.05.exe")
            time.sleep(10)
            fmod=open("script\csim.log","r")
            lineIn=fmod.readlines()
            for strIn in lineIn:
                strIn=strIn.strip("\n")
                if strIn.find("ERROR")>0:
                    print "\n"
                    print "executeATTScript failed."
                    fmod.close()
            fmod.close()
            filx=open("log.txt", "r")
            linez= filx.read()
            if str("fail: 1") in linez:
                sheet1.write(j,0,PLMN)
                sheet1.write(j,1,"FAIL to write IMSI")
                sheet1.write(j,2,"FAIL to write IMSI")
                sheet1.write(j,3,"FAIL to write IMSI")
                wb1.save('report.xls')

            elif str("pass") in linez:
                os.system("adb shell pm clear com.android.phone")
                os.system("adb reboot")
                time.sleep(85)
                from com.dtmilano.android.viewclient import ViewClient
                device, serialno=ViewClient.connectToDeviceOrExit()
                vc=ViewClient(device=device, serialno=serialno)
                if vc.findViewWithText("MAI TEST"):
                    vc.dump()
                    if vc.findViewById("android:id/button1"): 
                        vc.dump()
                        time.sleep(1)
                        vc.findViewById("android:id/button1").touch()
                        vc.dump()
                        print("clicked")
                        Network()
                        print("next func")
                        Call()    
                        SMS()   
                        wb1.save('report.xls')
                    else:
                        print("")

                else:
                    check()
                    os.system("adb shell cmd statusbar collapse")
                    if gets==str("4G Signal") or String==str("LTE Signal") or String==str("4G"):
                        refresh()
                        Call()    
                        SMS()   
                        wb1.save('report.xls')
                    else:
                        print(sendCommand("AT+CFUN=0"))
                        time.sleep(3)
                        print(sendCommand("AT+CFUN=1"))
                        time.sleep(3)
                        print(sendCommand("AT+CFUN=1,1"))
                        time.sleep(45)
                        Network()        
                        Call()    
                        SMS()   
                        wb1.save('report.xls')
                
               
            filx.close()
                
except:
    print("Exception occured........")
