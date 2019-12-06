
import os
import time
import re
import fnmatch
from xml.dom.minidom import parse
import xml.etree.ElementTree as ET
import xml.etree.ElementTree as ElementTree
import xml.dom.minidom
import serial
import serial.tools.list_ports
import subprocess
import glob
import time
import struct
import filecmp
import binascii
import win32com.client as win32
import struct
import Tkinter as tk
from ScrolledText import ScrolledText
import tkSimpleDialog
import tkMessageBox
from Tkinter import *
from serial.tools import list_ports
import shutil
from PIL import Image,ImageTk,ImageFilter
import subprocess
import logging
import os.path
import tkMessageBox
from xml.parsers.expat import ExpatError

# <The rest of your build script goes here.>


#Create and configure logger
LOG_FORMAT = "%(levelname)s %(asctime)s -%(message)s"
logging.basicConfig(filename="FINAL_RESULT.log",level=logging.DEBUG,format=LOG_FORMAT,filemode='w')
logger=logging.getLogger()

#Enter the NVs you want to skip begin
bannlist = ["73527","69689","66043","67225","1014","441","73833"]
#Enter the NVs you want to skip end

#Enter the NVs for which you want to use begin
EFSList = [66048,66472,70239]
#Enter the NVs for which you want to use end

negative_list=[]
fail_list=[]
failcount=0
oem_efs_list=[]
exception_count=0
exception_list=[]
cwd=os.getcwd()
port_num2=0

folderPlmnList=[]
ue_bannner=None
'''================================================================================================='''
'''================================================================================================='''

def isSIMdetected(Plmn):
    sim_plmn = os.popen("adb shell getprop gsm.sim.operator.numeric")
    sim_plmn = sim_plmn.read().strip()
    logger.info('#### Sim PLMN and Expected Plmn <Expected>:<'+Plmn+'>:<SIM>:<'+sim_plmn+'> ####')

    if(sim_plmn != ''):
        return True
    else:
        return False

def doSIMreset():
    os.popen("perl ResetPhone.pl  " + str(port_num2) + " " + str("00071") + " 0")
    time.sleep(3)

'''================================================================================================='''
def writePlmn(Plmn):
    mnc_len=2
    if Plmn.__len__() == 6:
        mnc_len = 3
        _mcc=Plmn[0:3]
        _mnc=Plmn[3:6]
    elif Plmn.__len__() == 5:
        mnc_len = 2
        _mcc = Plmn[0:3]
        _mnc = Plmn[3:5]

        logger.info("Entered MCC and MNC is " + _mcc + " " + _mnc)

    if mnc_len == 2:
        print "MNC 2 is Two Digit"
        command1 = 'com::send {at+csim=14,"00A40004026FAD";}\n'
        # command2='com::send {at+csim=10,"00B000000B";}\n'
        command3 = 'com::send {at+csim=18,"00D600000400000102";}\n'
        command4 = 'com::send {at+csim=14,"00A40004026F07";}\n'
        # command5='com::send {at+csim=10,"00B000000B";}\n'
        command6 = 'com::send {at+csim=28,"00D600000908'
        command6 += reverse('9' + _mcc + _mnc + '01')
        command6 += '00100111";}\n'
    elif mnc_len == 3:
        print "MNC 3 is Three Digit"
        command1 = 'com::send {at+csim=14,"00A40004026FAD";}\n'
        # command2='com::send {at+csim=10,"00B000000B";}\n'
        command3 = 'com::send {at+csim=18,"00D600000400000103";}\n'
        command4 = 'com::send {at+csim=14,"00A40004026F07";}\n'
        # command5='com::send {at+csim=10,"00B000000B";}\n'
        command6 = 'com::send {at+csim=28,"00D600000908'
        command6 += reverse('9' + _mcc + _mnc + '1')
        command6 += '00100111";}\n'

    commandList = []
    commandList.append('com::send {at+csim=14,"00A40004023F00";}\n')
    commandList.append('com::recv OK\n')
    commandList.append('com::send {at+csim=14,"00A40004027FF0";}\n')
    commandList.append('com::recv OK\n')
    commandList.append('com::send {at+csim=26,"0020000A083030303030303030";}\n')
    commandList.append('com::recv OK\n')
    commandList.append(command1)
    commandList.append('com::recv OK\n')
    # commandList.append(command2)
    # commandList.append('com::recv OK\n')
    commandList.append(command3)
    commandList.append('com::recv OK\n')
    commandList.append(command4)
    commandList.append('com::recv OK\n')
    # commandList.append(command5)
    # commandList.append('com::recv OK\n')
    commandList.append(command6)
    commandList.append('com::recv OK\n')

    file_object = open("script\csim.tcl", "w")
    file_object.writelines(commandList)
    file_object.close()

    os.system("ATTv1.05.exe")
    time.sleep(2)
    fmod = open("script\csim.log", "r")
    lineIn = fmod.readlines()
    for strIn in lineIn:
        strIn = strIn.strip("\n")
        if strIn.find("ERROR") > 0:
            print "\n"
            print "executeATTScript failed."
            fmod.close()
    fmod.close()
    return

def makeplmnlist(Plmnlist):
    mcc = 0
    mnc = 0
    i = 0
    plmn = []

    for plmns in Plmnlist:
        if (i % 2 == 0):
            mcc = plmns
            i += 1
        else:
            mnc = plmns
            if (len(mnc) == 1):
                mnc = '0' + mnc
            elif((len(mnc) == 1) and (mnc=='0')):
                mnc = "00"
            i += 1

        if (mcc and mnc != 0):
            plmn.append(mcc + mnc)
            mnc = 0
            mcc = 0
    return plmn

def getValueFromArray(root, efstype, key, NV, membername, nvname):
    # Read vendor.xml and list the MCCMNC for which volte,VoWIFI is enabled
    #print"NV:[" + NV + "]" + "[" + nvname + "]"
    for child in root.iter(efstype):
        if(child.attrib[key] == NV):
            for inner_elements in child.findall("Member"):
                if(inner_elements.text != None and inner_elements.attrib["name"].strip(" ") == membername):
                    #print "["+inner_elements.text.strip(" ")+"]"
                    return inner_elements.text.strip(" ")

def find_filename_gee(plmn_to_search):
    for row in folderPlmnList:
        folder_name = row[1]
        for plmn_list in row[0]:
            if plmn_to_search == plmn_list:
                logger.info("Mbn found with given Plmn.Folder name ["+folder_name+"] : Plmn ["+plmn_list+"]")
                return folder_name
    logger.error("No folder found for this Plmn : "+plmn_to_search)
    return None

def getPlmnData(root):
    Plmns = getValueFromArray(root, "NvTrlRecord", "category", "MCFG", "MCC_MNC_List", "Plmnlist")
    Plmns = Plmns.split(" ")
    if ( len(Plmns) == 1 and Plmns[0] == '0'):
        logging.fatal("<<ERROR>>No PLMN present for this operator")
        return None

    plmnlist = makeplmnlist(Plmns)
    if (len(plmnlist) == 0):
        logging.fatal("<<ERROR>>No PLMN present for this operator")
        return None

    return plmnlist

def locate(pattern, root):
    '''Locate all files matching supplied filename pattern in and below
    supplied root directory.'''
    for path, dirs, files in os.walk(os.path.abspath(root)):
        for filename in fnmatch.filter(files, pattern):
            yield os.path.join(path, filename)
'''================================================================================================='''
def listAllmbns():
    i=0
    for xml in locate("*.xml", os.getcwd()+'\\SourceMbnFiles\oem'):
        try:
            head, tail = os.path.split(xml)
            if (tail.startswith('mcfg_') and (not head.__contains__("\oem\CT")) and (not head.__contains__("TEST")) and (not head.__contains__("CMCC")) and (not head.__contains__("ROW")) and (not head.__contains__("\oem\CTA"))  and (not head.__contains__("\oem\CU"))):
                tree = ElementTree.parse(xml)
                root = tree.getroot()
                plmn = getPlmnData(root)
                #head, tail = os.path.split(head)
                if plmn != None:
                    list=plmn,xml
                    folderPlmnList.append(list)
                    print folderPlmnList
                i = i + 1
            else:
                logger.info("<<SKIP>>Skipped :"+xml)
                continue
        except (SyntaxError, ExpatError):
            logger.fatal("<<FATAL>>Xml format is wrong!")
            logger.fatal(xml)
    logger.info("total number of mbns ["+str(i)+"]")
    return str(i)
'''================================================================================================='''
def read_ue_value_with_efs(efs_path,bytes_to_read,bytes_to_skip):
    ###########################################READING EFS FILE AND COMPARING BEGIN ####################################################################################
    efs_name = get_efs_name(efs_path)

    os.popen("perl copy_phone_to_pc.pl  " + str(port2) + "  " + efs_path)
    shutil.copyfile('efs_buffer.txt', cwd + "\\" + "UE EFS" + "\\" + str(efs_name))
    if os.path.exists('efs_buffer.txt'):
        os.remove('efs_buffer.txt')
    f = open(cwd + "\\" + "UE EFS" + "\\" + str(efs_name), 'rb')
    f.seek(bytes_to_skip, 1)
    byte = f.read(bytes_to_read)
    if(bytes_to_read == 8):
        value = int(struct.unpack("<B", byte)[0])
    elif (bytes_to_read == 16):
        value = struct.unpack("<H", byte)[0]
    elif (bytes_to_read == 32):
        value = struct.unpack("<I", byte)[0]
    else:
        q=0
        value=byte
        string=""
        while (ord(value[q])):
            logger.debug("*********************************String was here*************************************************************" + "\n")
            string = string + value[q]
            q = q + 1
        value = string
    return value

def excel_result(checkstring,Nvs,parameters,QCvalues,UEvalues,results,ue_banner):
    p,l,m,n,o=2,2,2,2,2
    print "Making test report"
    logger.info("Making test report")
    output.insert(END,"\n Making test report")
    cwd = os.getcwd()
    #os.popen("taskkill /im EXCEL.exe /f")
    excel2 = win32.gencache.EnsureDispatch('Excel.Application')
    wb_result = excel2.Workbooks.Open(cwd + "\\Result_template.xlsx")
    ws_result = wb_result.Worksheets(1)
    output.see("end")
    main.update_idletasks()
    for x in Nvs:
        ws_result.Cells(p,1).Value=str(x)
        p=p+1
    for x in parameters:
        ws_result.Cells(l,2).Value=str(x)
        l=l+1
    for x in QCvalues:
        ws_result.Cells(m,3).Value=str(x)
        m=m+1
    for x in UEvalues:
        ws_result.Cells(n,4).Value=str(x)
        n=n+1
    for x in results:
        ws_result.Cells(o,5).Value=str(x)
        o=o+1

    excel2.DisplayAlerts = False
    ws_result.SaveAs(os.getcwd() +"\\Results\\"+str(ue_banner)+" "+str(checkstring)+".xlsx")
    excel2.Application.Quit()
    print "Test report saved"
    logger.info("Test report saved")
    output.see("end")
    output.insert(END,"\n\nReport created")

    output.update_idletasks()

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

def plmn_list_find_parent(checkstring):
    main.update()
    print "Currently in plmn_list_find_parent function"
    logger.info("Currently in plmn_list_find_parent function")
    excel6 = win32.gencache.EnsureDispatch('Excel.Application')
    wbn_data = excel6.Workbooks.Open(cwd + "\\PLMN_list.xls")
    wg = wbn_data.Worksheets(1)
    i=1
    while(wg.Cells(i,1).Value!=None):
        logger.info("Currently in row i :"+str(i))
        j=1
        while(wg.Cells(i,j).Value!=None):
            if(str(wg.Cells(i,j).Value)==str(checkstring)):
                logger.info("The parent mcc mnc is :"+str(wg.Cells(i,1).Value))
                return str(wg.Cells(i,1).Value)
            j=j+1
        i=i+1
    logger.critical("The mcc you mentioned is not found in PLMN_list or in in mbn folder.")
    logger.critical("Please try again after checking")
    excel6.Application.Quit()

def read_ue_value_with_no_efs_new(index):
        print "Currently in read_ue_value_with_no_efs_new function"
        logger.info("Currently in read_ue_value_with_no_efs_new function")
        main.update()
        main.update_idletasks()
        fmod = open("NV_buffer.txt", "r")
        lines = fmod.read().split(',')
        count=len(lines)
        k=index+2
        if k>(count-1):
            logger.info("The index you are reading is not there in the list please check using qxdm")
            return "Novalue"
        ue_value = lines[k]
        logger.info("Value of i is "+str(index))
        logger.info("Value of k is "+str(k))
        logger.info("The ue_value is "+str(ue_value))
        return ue_value

def readNVItemData_new(NVs,parameters,QCvalues,UEvalues,results,tree):
    print ("Currently in readNVItemData_new function\n")
    logger.info("Currently in readNVItemData_new function\n")
    root = tree.getroot()
    NvItemDatas = root.findall('NvItemData')
    for NvItemData in NvItemDatas:
        NV= NvItemData.attrib['id']
        if NV in bannlist:
            logger.info("This NV is there in bannlist")
            continue
        logger.info("NV ="+str(NV))

        ##########Pull the NV file begin###################
        fcre = open("NV_buffer.txt", "w")
        fcre.close()
        os.popen("perl QXDMClientRequestNVRead.pl  " + str(port_num2) +" "+str(NV)  + " 0")
        ##########Pull the NV file end######################

        index=0
        Members = NvItemData.findall('Member')
        for Member in Members:
            parameter= Member.attrib['name']
            if parameter=="ReservedBytes" or parameter=="reserved":
                logger.info("This parameter is reserved bytes skipping")
                index = index + 1
                continue
            #print Member.text
            QCvalue= Member.text

            if QCvalue=="None" or QCvalue == None:
                logger.info("This NV has no QCvalue")
                index = index + 1
                continue


            UEvalue=read_ue_value_with_no_efs_new(index)

            if UEvalue in thisdict:
                logger.info("this value is there in dictionary")
                UEvalue=thisdict[UEvalue]

            index = index + 1

            if UEvalue=="Novalue":
                logger.info("No value in UE for this NV skipping")
                continue

            try:
                if UEvalue[0]=="-":
                    logger.info("This is a negative value skipping")
                    continue

                if UEvalue[1]=="x":
                    logger.info("This is a hex value..Changing hex value to int")
                    UEvalue=UEvalue[2:]
                    UEvalue=int(UEvalue,16)
                    UEvalue=str(UEvalue)
            except:
                logger.warning("There is no index 1 in the string you are checking")

            parameters.append(parameter)
            QCvalues.append(QCvalue)
            logger.info("QCValue is " + str(QCvalue))
            NVs.append(NV)
            UEvalues.append(UEvalue)
            if str(QCvalue).strip()==str(UEvalue).strip():
                results.append("SUCCESS")
                logger.info("Success")
            else:
                results.append("FAIL")
                logger.info("FAIL")

def readNvEfsItemData_new(NVs, parameters, QCvalues, UEvalues, results,tree):
    print ("Currently in readNvEfsItemData_new function\n")
    logger.info("Currently in readNvEfsItemData_new function\n")
    root = tree.getroot()
    NvEfsItemDatas= root.findall('NvEfsItemData')
    for NvEfsItemData in NvEfsItemDatas:
        logger.info("NVEFSItemData = "+str(NvEfsItemData))
        mcfgAttributes=NvEfsItemData.attrib['mcfgAttributes']
        if mcfgAttributes!="0x19":
            continue
        NV= NvEfsItemData.attrib['id']
        if NV in bannlist:
            continue
        if NV in EFSList:
            logger.info("this NV will be read by decoding efs files")
            readNvEfsItemData(NVs, parameters, QCvalues, UEvalues, results,tree)

        ##############Pull the NV file begin############
        fcre = open("NV_buffer.txt", "w")
        fcre.close()
        os.popen("perl QXDMClientRequestNVRead.pl  " + str(port_num2) +" "+str(NV)  + " 0")
        ##############Pull the NV file end############

        Members = NvEfsItemData.findall('Member')
        index=0
        for Member in Members:
            logger.info("Member name ="+str(Member.attrib['name']))
            logger.info("NV = "+str(NV))
            parameter= Member.attrib['name']
            if parameter=="ReservedBytes" or parameter=="reserved":
                index = index + 1
                continue
            #print Member.text
            QCvalue= Member.text

            if QCvalue=="None" or QCvalue == None:
                logger.info("QC value is None...skipping this NV")
                index = index + 1
                continue


            UEvalue=read_ue_value_with_no_efs_new(index)
            if UEvalue in thisdict:
                logger.info("this value is there in dictionary")
                UEvalue=thisdict[UEvalue]
            index = index + 1

            if UEvalue=="Novalue":
                logger.debug("No value in UE for this NV skipping")
                continue

            try:
                if UEvalue[0]=="-":
                    logger.debug("This is a negative value skipping")
                    continue

                if UEvalue[1]=="x":
                    logger.debug("This is a hex value..Changing hex value to int")
                    UEvalue=UEvalue[2:]
                    UEvalue=int(UEvalue,16)
                    UEvalue=str(UEvalue)
            except:
                logger.warning("There is no UEvalue[1]")

            parameters.append(parameter)
            QCvalues.append(QCvalue)
            logger.info("QCValue is " + str(QCvalue))

            UEvalues.append(UEvalue)
            NVs.append(NV)
            if str(QCvalue).strip()==str(UEvalue).strip():
                results.append("SUCCESS")
                logger.info("Success\n")
            else:
                results.append("FAIL")
                logger.info("Fail\n")

def get_mcc(checkstring):
    logger.debug("Currently in get_mcc function\n")
    mcc=checkstring[:3]
    return mcc

def get_mnc(checkstring):
    logger.debug("Currently in get_mnc function\n\n")
    mnc=checkstring[3:]
    return mnc

def get_ue_banner():
    try:
        print "Currently in get_ue_banner function"
        #subprocess.Popen("adb shell setprop persist.sys.usb.config manufacture,adb")
        output.insert(END,"\n\nPrinting the banner....\n")
        main.update()
        main.update_idletasks()
        logger.debug("Currently in get_ue_banner function\n")
        fcre = open("NV_buffer.txt", "w")
        fcre.close()
        os.popen("perl QXDMClientRequestNVRead.pl  " + str(port_num2) + "  00071  0")
        fmod = open("NV_buffer.txt", "r")
        lines = fmod.read().split(',')
        ue_bannner = lines[2]
        print "The Banner in UE is:",ue_bannner
        logger.info("The Banner in UE is:"+str(ue_bannner))
        output.insert(END,"\n\n"+ue_bannner+"\n")
        fmod.close()
        return ue_bannner
    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        logger.critical("This is an exception" + str(message) + "\n\n")
        tkMessageBox.showinfo("Error", str(message))

def find_filename(arr,checkstring):
    checkstring=plmn_list_find_parent(checkstring)
    logger.debug("Currently in find_filename function\n")
    for filename in arr:
        if str(checkstring) in str(filename):
            logger.info("The parent folder for this mcc mnc is"+str(filename))
            return filename
    print "Wrong mcc mnc or the operator is not listed in the folder"
    output.insert(END, "\n\nWrong mcc mnc or the operator is not listed in the folder\n")
    logger.error("Wrong mcc mnc or the operator is not listed in the folder \n")
    t=1
    while t!=10:
        t=t+1
        main.update()
        time.sleep(1)
    exit()

def get_efs_name(efs_path):
    logger.debug("Currently in get_efs_name function\n")
    rev_efs_path=""
    rev_efs_path = efs_path[::-1]
    k = 0
    efs_filename = ""
    while (rev_efs_path[k] != '/'):
        efs_filename = efs_filename + rev_efs_path[k]
        k = k + 1
    efs_filename = efs_filename[::-1]
    logger.info("The get_efs_name is working on:"+str(efs_filename)+"\n")
    return efs_filename

def get_non_efs_name(non_efs_path):
    logger.debug("Currently in get_non_efs_name function\n")
    non_efs_name=os.path.basename(non_efs_path)
    return non_efs_name

def get_efs_files_path(non_efs_path):
    logger.debug("Currently in get_efs_files_path function\n")
    efs_files_path=non_efs_path[37:]
    return efs_files_path

def set_mcc_mnc1(mcc,mnc):
    output.insert(END, "\n\nWriting MCC MNC......" + "\n")
    main.update()
    main.update_idletasks()
    logger.debug("Writing MCC MNC......" + "\n")
    fmod = open("att.ini", "r")
    line = fmod.read()
    fmod.close()
    line1 = re.sub('com=.*', "com=" + str(port_num1), line)
    fmod = open("att.ini", "w")
    fmod.write(line1)
    fmod.close()

    t=1
    while t!=10:
        time.sleep(1)
        main.update()
        t=t+1
    #os.system("python runMccMnc.py 1 " + str(mcc) + " " + str(mnc))

    #################TESTING#########################################

    print "\nEntered MCC and MNC is " + mcc + " " + mnc

    command1 = ""
    command2 = ""
    command3 = ""
    command4 = ""
    command5 = ""
    command6 = ""
    if mnc.__len__() == 2:
        print "MNC 2 is Two Digit"
        command1 = 'com::send {at+csim=14,"00A40004026FAD";}\n'
        # command2='com::send {at+csim=10,"00B000000B";}\n'
        command3 = 'com::send {at+csim=18,"00D600000400000102";}\n'
        command4 = 'com::send {at+csim=14,"00A40004026F07";}\n'
        # command5='com::send {at+csim=10,"00B000000B";}\n'
        command6 = 'com::send {at+csim=28,"00D600000908'
        command6 += reverse('9' + mcc + mnc + '01')
        command6 += '00100111";}\n'
    elif mnc.__len__() == 3:
        print "MNC 3 is Three Digit"
        command1 = 'com::send {at+csim=14,"00A40004026FAD";}\n'
        # command2='com::send {at+csim=10,"00B000000B";}\n'
        command3 = 'com::send {at+csim=18,"00D600000400000103";}\n'
        command4 = 'com::send {at+csim=14,"00A40004026F07";}\n'
        # command5='com::send {at+csim=10,"00B000000B";}\n'
        command6 = 'com::send {at+csim=28,"00D600000908'
        command6 += reverse('9' + mcc + mnc + '1')
        command6 += '00100111";}\n'


    commandList = []
    commandList.append('com::send {at+csim=14,"00A40004023F00";}\n')
    commandList.append('com::recv OK\n')
    commandList.append('com::send {at+csim=14,"00A40004027FF0";}\n')
    commandList.append('com::recv OK\n')
    commandList.append('com::send {at+csim=26,"0020000A083030303030303030";}\n')
    commandList.append('com::recv OK\n')
    commandList.append(command1)
    commandList.append('com::recv OK\n')
    # commandList.append(command2)
    # commandList.append('com::recv OK\n')
    commandList.append(command3)
    commandList.append('com::recv OK\n')
    commandList.append(command4)
    commandList.append('com::recv OK\n')
    # commandList.append(command5)
    # commandList.append('com::recv OK\n')
    commandList.append(command6)
    commandList.append('com::recv OK\n')

    main.update()
    file_object = open("script\csim.tcl", "w")
    file_object.writelines(commandList)
    file_object.close()

    subprocess.Popen("ATTv1.05.exe")
    t=1
    while t!=5:
        time.sleep(1)
        t=t+1
        main.update()
    fmod = open("script\csim.log", "r")
    lineIn = fmod.readlines()
    for strIn in lineIn:
        strIn = strIn.strip("\n")
        if strIn.find("ERROR") > 0:
            print "\n"
            print "executeATTScript failed."
            fmod.close()
    fmod.close()


    ################TESTING END#########################################




    ######checking mcc mnc writing error#######
    fmod = open("script\csim.log", "r")
    lineIn = fmod.readlines()
    for strIn in lineIn:
        strIn = strIn.strip("\n")
        if strIn.find("ERROR") > 0:
            print "\n"
            print "Writing MCCMNC failed."
            logger.error("Writing MCCMNC fail" + "\n")
            output.insert(END, "Writing MCCMNC fail" + "\n")
            fmod.close()
            exit()
            ######checking mcc mnc writing error#######
    t=1
    while t!=5:
        time.sleep(1)
        main.update()
        t=t+1
    main.update()
    subprocess.Popen("adb reboot")
    t=1
    while t!=70:
        time.sleep(1)
        main.update()
        t=t+1
    main.update()
    subprocess.Popen("adb shell setprop persist.sys.usb.config manufacture,adb")
    t=1
    while t!=10:
        time.sleep(1)
        main.update()
        t=t+1
    main.update()
    logger.info("Writing MCCMNC Success."+"\n")
    output.insert(END,"Writing MCCMNC Success."+"\n")
    main.update_idletasks()

def get_xml_banner(tree):
    print("Currently in get_xml_banner function\n")
    logger.info("Currently in get_xml_banner function\n")
    root = tree.getroot()

    NvItemDatas = root.findall('NvItemData')
    logger.info("NvItemDatas = "+str(NvItemDatas))
    for NvItemData in NvItemDatas:
        logger.info(NvItemData.attrib['id'])
        if NvItemData.attrib['id'] == str(71):
            Members = NvItemData.findall('Member')
            for Member in Members:
                logger.info("Name of Member="+str(Member.text))
                logger.info("Text of Member="+str(Member.attrib['name']))
                banner_in_xml=Member.text
                logger.debug("The banner in xml is :" + str(banner_in_xml))
                return banner_in_xml

def check_for_mcc_mnc_new_chipset(checkstring):
    print "Currently in check_for_mcc_mnc_new_chipset function"
    logger.info("Currently in check_for_mcc_mnc_new_chipset function")
    #checkstring = tkSimpleDialog.askstring("MCC-MNC", "Enter MCC MNC. For ex:- MCC_MNC")
    NVs = []
    parameters = []
    QCvalues = []
    UEvalues = []
    results = []

    main.update()
    output.insert(END,"\n\nChecking for :"+str(checkstring)+"......\n\n")
    logger.info("\n\nChecking for :"+str(checkstring)+"......\n\n")
    main.update_idletasks()
    output.see("end")
    global exception_count
    global exception_list
    global negative_list
    global fail_list
    global failcount
    global oem_efs_list
    exception_count = 0
    exception_list = []
    negative_list = []
    fail_list = []
    failcount = 0
    oem_efs_list = []
    writePlmn(checkstring)
    doSIMreset()
    current = os.getcwd()+'\\SourceMbnFiles\oem'

    selected_folder_name = find_filename_gee(checkstring)
    logger.info("Selected folder name is : "+str(selected_folder_name))

    if selected_folder_name == None:
        output.insert(END, "Mbn file for this PLMN not present in the Source MbnFile\n"+checkstring)
        return

    #tree = ET.parse(current + "\\sdm\\generic\\oem\\" + selected_folder_name + "\\mcfg_sw_gen_Commercial.xml")
    tree = ET.parse(selected_folder_name)

    while (1):
        if (isSIMdetected(checkstring) == True):
            logging.info("Sim detected")
            break
        else:
            logging.info("Sim not detected...")
            time.sleep(1)
            continue
    ue_banner = get_ue_banner()
    xml_banner= get_xml_banner(tree)
    if(ue_banner != xml_banner):
        logger.error("The banner in UE is not the same as in the XML you are comparing it with\n")
        output.insert(END,"The banner in UE is not the same as in the XML you are comparing it with\n")
        logger.error("Skipping it and going to check new PLMN")
        return

    logger.info("The banner in UE is the same as in the XML you are comparing it with\n")
    output.insert(END,"\nThe banner in UE is the same as in the XML you are comparing it with\n")
    output.insert(END,"\nNow starting comparison ....\n")

    start=time.time()
    # Read and compare NVItemData :- This will be NVs without a corresponding EFS file in code or device
    readNVItemData_new(NVs,parameters,QCvalues,UEvalues,results,tree)
    end=time.time()
    logger.info("Time taken for readNVItemData_new function is:"+str(end - start))

    start = time.time()
    # Read and compare NVEfsItemData :- This will be NVs with a corresponding EFS file in device but not in modem code
    readNvEfsItemData_new(NVs, parameters, QCvalues, UEvalues, results,tree)
    end=time.time()
    logger.info("Time taken for readNvEfsItemData_new function is:"+str(end - start))


    # Read anc compare NvEfsFile :- This will be Profiles/DataFiles with corresponding EFS file in both device and modem code
    ###readNvEfsFile(NVs, parameters, QCvalues, UEvalues, results,tree)


    excel_result(checkstring, NVs, parameters, QCvalues, UEvalues, results, ue_banner)
    output.see("end")

def readNvEfsItemData(NVs,parameters,QCvalues,UEvalues,results,tree):
    logger.debug("Currently in readNvEfsItemData function\n")
    print("Currently in readNvEfsItemData function\n")
    root = tree.getroot()
    print root
    NvEfsItemDatas = root.findall('NvEfsItemData')
    print NvEfsItemDatas
    for NvEfsItemData in NvEfsItemDatas:
        bytes_to_skip=0
        NV = NvEfsItemData.attrib['id']
        efs_path=NvEfsItemData.attrib['fullpathname']

        Members = NvEfsItemData.getiterator()
        for Member in Members:
            logger.info("Member name ="+str(Member.attrib['name']))
            parameter = Member.attrib['name']
            parameters = parameters.append(parameter)
            # print Member.text
            UEvalue = Member.text
            UEvalues.append(UEvalue)
            typel=Member.attrib['type']
            sizeOf=Member.attrib['sizeOf']
            real_type=typel[4:]
            if typel=="string":
                bytes_to_read=int(sizeOf)
            else:
                bytes_to_read=int(real_type)*int(sizeOf)
            QCvalue = read_ue_value_with_efs(efs_path,bytes_to_read,bytes_to_skip)
            bytes_to_skip=bytes_to_skip+bytes_to_read
            NVs.append(NV)
            QCvalues.append(QCvalue)
            if QCvalue == UEvalue:
                results = results.append("SUCCESS")
            else:
                results = results.append("FAIL")

def plmn_list_run():
    #os.popen("taskkill /im EXCEL.exe /f")
    subprocess.Popen("adb shell setprop persist.sys.usb.config manufacture,adb")
    check_plmn_list=[]
    main.update()
    listAllmbns()
    print "Currently in plmn_list_find_parent function"
    logger.info("Currently in plmn_list_find_parent function")
    excel7 = win32.gencache.EnsureDispatch('Excel.Application')
    wbn_data = excel7.Workbooks.Open(cwd + "\\PLMN_list.xls")
    wg = wbn_data.Worksheets(1)
    i=1
    while(wg.Cells(i,1).Value!=None):
        j=1
        while (wg.Cells(i,j).Value!=None):
            check_plmn_list.append(str(int(wg.Cells(i,j).Value)))
            j=j+1
        i=i+1
    excel7.Application.Quit()
    for plmn in check_plmn_list:
        print "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$Testing new Plmn:"+ plmn+"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$Testing"
        logger.info("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$Testing new Plmn:"+ plmn+"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$Testing")
        ue_bannner=None
        check_for_mcc_mnc_new_chipset(plmn)

def ports_print():
    ports = list(serial.tools.list_ports.comports())
    output.insert(END,"\n\nPorts connected are :-\n")
    for p in ports:
        output.insert(END,"\n" + str(p[1]))

#########################################################################

ports = list(serial.tools.list_ports.comports())

for p in ports:
    logger.info("List of ports are:\n"+str(p))

try:
    cdc1 = next(list_ports.grep("Android Adapter PCUI"))
    port1 = cdc1[0]
    port_num1 = port1[3:]
    logger.info("port for mcc mnc is "+port_num1+"\n")
except StopIteration:
    logger.error("\\\\\\\\\\\\\\\No DEVICE found\\\\\\\\\\n")

try:
    cdc2 = next(list_ports.grep("DBAdapter Reserved Interface"))
    port2 = cdc2[0]
    port_num2 = port2[3:]
    logger.info("port for mcc mnc is "+port_num2+"\n")
    # port_num=cdc[0]
except StopIteration:
    logger.error("\\\\\\\\\\\\\\\No DEVICE found\\\\\\\\\\n")

thisdict =	{
    "UE_USAGE_SETTING_VOICE_CENTRIC": "0",
    "UE_USAGE_SETTING_DATA_CENTRIC": "1",
    "SYS_VOICE_DOMAIN_PREF_CS_VOICE_ONLY": "0",
    "SYS_VOICE_DOMAIN_PREF_IMS_PS_VOICE_ONLY":"1",
    "SYS_VOICE_DOMAIN_PREF_CS_VOICE_PREFFERED":"2",
    "SYS_VOICE_DOMAIN_PREF_IMS_PS_VOICE_PREFFERED":"3",
    "SYS_SMS_DOMAIN_PREF_NONE":"-1",
    "SYS_SMS_DOMAIN_PREF_PS_SMS_NOT_ALLOWED":"0",
    "SYS_SMS_DOMAIN_PREF_PS_SMS_PREF":"1",
    "SYS_SMS_DOMAIN_PREF_MAX":"2",
    "SYS_PS_SUPP_DOMAIN_PREF_AUTO":"0",
    "SYS_PS_SUPP_DOMAIN_PREF_CS_ONLY":"1",
    "SYS_PS_SUPP_DOMAIN_PREF_PS_ONLY":"2",
    "SYS_PS_SUPP_DOMAIN_PREF_PS_PREF":"3",
    "QPE_ENABLE_REREG_ON2G3G_INVALID":"0",
    "QPE_ENABLE_REREG_ON2G3G_ON":"1",
    "QPE_ENABLE_REREG_ON2G3G_OFF":"2",
    "IMS_SILENT_REDIAL_DISABLE":"0",
    "IMS_SILENT_REDIAL_ENABLE":"1",
    "IMS_E911_TEST_MODE_OFF":"0",
    "IMS_E911_TEST_MODE_ON":"1",
    "eQP_IMS_XCAP_GBA_MODE_NONE":"0",
    "eQP_IMS_XCAP_GBA_ALWAYS":"1",
    "eQP_IMS_XCAP_GBA_ON_REQ":"2",
    "FALSE":"0",
    "false":"0",
    "TRUE":"1",
    "true":"1",
    "eQP_IMS_XCAP_GBA_TLS_NONE":"0",
    "eQP_IMS_XCAP_GBA_TLS_SK_CB":"1",
    "eQP_IMS_XCAP_GBA_TLS_ON_DEMAND":"2",
    "UT_MEDIA_ELEMENT_NONE":"0",
    "UT_AUDIO_WITH_AUDIO_TAG_AND_VIDEO_WITH_ONLY_VIDEO_TAG":"1",
    "UT_AUDIO_WITH_AUDIO_TAG_AND_VIDEO_WITH_AUDIO_AND_VIDEO_TAGS":"2",
    "UT_AUDIO_WITHOUT_MEDIA_TAG_AND_VIDEO_WITH_VIDEO_TAG":"3",
    "UT_EMPTY_SIB_AS_NONE":"0",
    "UT_EMPTY_SIB_AS_REJECT":"1",
    "UT_EMPTY_SIB_AS_AUDIO":"2",
    "UT_EMPTY_SIB_AS_AUDIO_AND_VIDEO":"3",
    "NV_UIM_FIRST_INST_CLASS_GSM_SIM":"0",
    "NV_UIM_FIRST_INST_CLASS_UMTS_SIM":"1",
    "NV_UIM_FIRST_INST_CLASS_USB_UICC":"2",
    "NV_UIM_FIRST_INST_CLASS_USB_UICC_RST_HIGH":"3",
    "S/W":"0",
    "USIM":"1",
    "DISABLED":"0",
    "ENABLED":"1",
    "Determine Mode Automatically":"4"
}
try:
    main=Tk()
    main.title("MBN CHECKING")
    main.geometry("600x800")

    icon=PhotoImage(file="icon5.gif")
    main.tk.call("wm",'iconphoto',main._w,icon)

    output= ScrolledText(main,width=82,height=50,background="white",undo=True)
    output.pack(expand=True,fill='both')
    output.place(x=0,y=100)

    b9 = Button(main, width=15, text="Run from PLMN list", command=plmn_list_run)
    b9.place(x=20, y=40)
    b9.config(relief=RAISED)


    main.mainloop()

except Exception as ex:
    template = "An exception of type {0} occurred. Arguments:\n{1!r}"
    message = template.format(type(ex).__name__, ex.args)
    print message
    logger.critical("This is an exception" + str(message)+ "\n\n")
