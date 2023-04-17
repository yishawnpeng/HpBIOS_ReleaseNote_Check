############################
#Shawn.Peng@quantatw.com
#Can help you to double confirm info between BCU and release_note
#support relase_note is .xlsm not .docx(usually before G5)
############################
#Version List
#v.1 support AMD
############################
#pip3 install pandas re os openpyel pypiwin32 shutil

import pandas as pd     #excel
#import re               #regular expression    #imported from lib#
#import os               #dir                   #imported from lib#
import sys              #exit don't use os.exit
#import shutil           #move file
#from datetime import date
from win32com.client import * # GetFileVersion from exe

from lib import *

#printDebug = 1
#writeSuccess=[]
#writeFail=[]
#errorList=[]
AMDPlatformDict = {"R24","R26","S25","S27","S29","T25","T26","T27"}
isAMDAMDPlatform = None

#get FUR.exe info
information_parser = Dispatch("Scripting.FileSystemObject")

#Get Release Note File
allDir = os.listdir( os.getcwd() ) #list
excelName = re.compile("\w.*Release_Note_\d*\.xlsm") #{not~}{any}Release_Note_{number}.xlsm
excelName = list( filter( excelName.match, allDir ) )
if not excelName : #empty
    print("Release_Note_excel is not exist or being opened!")
    print("Format should be {any}Release_Note_{number}.xlsm !")
    os.system("pause")
    sys.exit()
else :
    rName = excelName[0]
    platform = rName.split("_")[2]
    version = rName.split("_")[-1].split(".")[0]
    #print(version)
    print( "Choose release note: " + rName )
    #Check platfrom
    if ( platform in AMDPlatformDict ):
        amdz_content=[]
        isAMDAMDPlatform = True
        getAMDZInfo(amdz_content)
    else :
        isAMDAMDPlatform = False
    #Get item name and info of this time
    try :
        #Item name
        rRowInfoName = pd.read_excel( rName, sheet_name = "AMDPlatformHistory", usecols=[0] ) if isAMDAMDPlatform \
                else pd.read_excel( rName, sheet_name = "IntelPlatformHistory", usecols=[0] ) 
        #This time
        rRowData = pd.read_excel( rName, sheet_name = "AMDPlatformHistory", usecols=[1] ) if isAMDAMDPlatform \
                else pd.read_excel( rName, sheet_name = "IntelPlatformHistory", usecols=[1] ) 
        rRowInfoName = rRowInfoName[rRowInfoName.columns[0]].tolist()
        #Find Item Range
        startIndex = rRowInfoName.index("System BIOS Version")
        endIndex = rRowInfoName.index("Sprint Release Note")
    except Exception :
        #print(Exception)
        print("Get release note info! May be ceil name error.")
        os.system("pause")
        sys.exit()

#Create Excel
try:
    writer = pd.ExcelWriter("result_RN.xlsx")
    outputFile = []
    outputFile_PlatformHistory = pd.DataFrame( index = rRowInfoName[startIndex:endIndex], \
                                              columns = ["ThisTimeUpdate", "Geted", "Result"] )
    outputFile.append(outputFile_PlatformHistory) #Sheet No.1
except Exception:
    #print(Exception)
    print("Creat excel fail or being opened!")
    os.system("pause")
    sys.exit()

#####Write Release Note Info
#version = rRowData[rRowData.columns[0]].tolist()[startIndex]
#This time info
outputFile[0].iloc[:, 0] = rRowData[rRowData.columns[0]].tolist()[startIndex:endIndex]
#print(rRowData[rRowData.columns[0]].tolist()[startIndex:endIndex])

#Get BCU File
bcu_name = re.compile(".*BCU\.txt") # {any}BCU.txt
bcu_name = list( filter( bcu_name.match, allDir ) )
if not bcu_name :   #Try .*bcu.txt
    bcu_name = re.compile(".*bcu\.txt") # {any}bcu.txt
    bcu_name = list( filter( bcu_name.match, allDir ) )
    
if not bcu_name :   #empty
    print("You don't have BCU file or being opened!")
    print("Format should be {any}BCU.txt or {any}bcu.txt !")
    os.system("pause")
    sys.exit()
else :
    print("Choose BCU :" + bcu_name[0])
    bcu_content=[]
    with open(bcu_name[0]) as f:
        for line in f.readlines():
            bcu_content.append(line)
    

#Find same in release note
rRowInfoName = rRowInfoName[startIndex:endIndex]
#print(bcu_content[ bcu_content.index(rRowInfoName[0]+"\n")+1])
for i in rRowInfoName:
    if type(i) != str : # skip nan
        continue
    try :
        name = i +"\n"
        index = bcu_content.index(name)+1
        if i == "System BIOS Version" :
            temp = bcu_content[index].split()[2]
        elif i =="BIOS Build Version" :
            temp = bcu_content[index].strip()
        else :
            temp = bcu_content[index]
        #print(temp)##########
        outputFile[0].at[i, "Geted"] = temp
    except Exception :
        if i == "Build Date" :
            try:
                bdate = bcu_content[bcu_content.index("System BIOS Version\n")+1].split()[-1]
                #print(bdate)##########
                outputFile[0].at[i, "Geted"] = bdate
                continue
            except :
                pass
        elif i == "CHECKSUM":
            binFile = ""
            #G5 check
            if isAMDAMDPlatform and os.path.isfile(".\\Global\\BIOS\\" + platform + "_" + version + ".bin") :
                binFile = ".\\Global\\BIOS\\" + platform + "_" + version + ".bin"
            #G5 check
            elif not isAMDAMDPlatform and os.path.isfile(".\\Global\\BIOS\\" + platform + "_" + version + "_32.bin") :
                binFile = ".\\Global\\BIOS\\" + platform + "_" + version + "_32.bin"

            if binFile :
                with open(binFile, 'rb') as f:
                    content = f.read()
                    binary_sum = sum(bytearray(content))
                    binary_sum = hex(binary_sum & 0xFFFFFFFF)
                    f.close()
                #x need lower, other need upper
                binary_sum = "0x"+binary_sum.split("x")[-1].upper()
                #binary_sum=binary_sum.upper()
                #print(binary_sum)
                outputFile[0].at[i, "Geted"] = binary_sum
            else :
                print("\nCan note find biniary : "+platform + "_" + version + ".bin")
                print("Format : {Platform}_{version}.bin (Fist letter upper) ")
            continue
        elif i == "Sprint" :
            print("\nSprint info in local build BCU!")
        elif i == "AMD Agesa PI" :
            if amdz_content :
                agesaPI = re.compile("AGESA:.*")
                agesaPI = list( filter( agesaPI.match, amdz_content ) )
                agesaPI = agesaPI[0].split()[-1] #agesaPI = agesaPI[0].split()[-1]
                agesaPI = "ComboAm4PI " + agesaPI
                #print(agesaPI)
                outputFile[0].at[i, "Geted"] = agesaPI
                continue
            else :
                print("\nCan not find PI because not have amdz!")
        elif i == "SMU FW" : #mayby dec to hex
            if amdz_content :
                PSPandSMU = re.compile("SMU:.*")
                PSPandSMU = list( filter( PSPandSMU.match, amdz_content ) )
                PSPandSMU = PSPandSMU[0].split()
                SMUFW = PSPandSMU[1].split("(")[0]
                #print(SMUFW)
                outputFile[0].at[i, "Geted"] = SMUFW
                continue
            else :
                print("\nCan not find SMU FW because not have amdz!")
        elif i == "PSP FW" :
            if amdz_content :
                PSPandSMU = re.compile("SMU:.*")
                PSPandSMU = list( filter( PSPandSMU.match, amdz_content ) )
                PSPandSMU = PSPandSMU[0].split()
                PSPFW = PSPandSMU[2]
                if "(" in PSPFW  :
                    PSPFW = PSPFW.split("(")[1][:-1]
                else :
                    PSPFW = PSPFW[2:]
                realPSPFW = ""
                for j in range(0,len(PSPFW),2) :
                    if PSPFW[j] == "0" and PSPFW[j+1] == "0" :
                        realPSPFW = realPSPFW + "0."
                    elif PSPFW[j] == "0" and PSPFW[j+1] != "0" :
                        realPSPFW = realPSPFW + PSPFW[j+1] + "."
                    else : 
                        realPSPFW = realPSPFW + PSPFW[j] + PSPFW[j+1] + "."
                realPSPFW = realPSPFW[:-1]
                outputFile[0].at[i, "Geted"] = realPSPFW
                #print(realPSPFW)
                continue
            else :
                print("\nCan not find PSP FW because not have amdz!")
        elif i == "EC/SIO F/W" :
            try:
                sio = bcu_content[bcu_content.index("Super I/O Firmware Version\n")+1].split()[-1]
                #print(sio)
                outputFile[0].at[i, "Geted"] = sio
                continue
            except :
                pass
        elif isTypecPD(i) :
            try :
                firstPD = bcu_content[bcu_content.index("USB Type-C Controller(s) Firmware Version:\n")+1].split()[-1]
                secondPD = bcu_content[bcu_content.index("USB Type-C Controller(s) Firmware Version:\n")+2].split()[-1]
                numdot=re.compile("[0-9]+")
                if not numdot.match(firstPD) :
                    outputFile[0].at[i, "Geted"] = "N\A"
                #check secondPD is exit or not
                elif numdot.match(firstPD) and not numdot.match(secondPD) :
                    print(firstPD)
                    outputFile[0].at[i, "Geted"] = firstPD
                    continue
                else :
                    print(firstPD)
                    print(secondPD)
                    outputFile[0].at[i, "Geted"] = firstPD + "\n" + secondPD
                    continue
            except :
                pass
        elif i == "AMD Legacy VBIOS" :
            if amdz_content :
                vBIOS = re.compile("VBIOS Info.*")
                vBIOS = list( filter( vBIOS.match, amdz_content ) )
                vBIOS = vBIOS[0].split()[3][0:-1]
                #print(vBIOS)
                outputFile[0].at[i, "Geted"] = vBIOS
                continue
            else :
                print("\nCan not find SMU FW because not have amdz!")
        elif i == "AMD GOP EFI Driver" :
            try :
                gOP = re.compile("Rev.*")
                gOP = list( filter( gOP.match, bcu_content[bcu_content.index("Video BIOS Version\n")+1].split() ) )
                gOP = gOP[0][4:-4]
                #print(gOP)
                outputFile[0].at[i, "Geted"] = gOP
                continue
            except Exception:
                print("AMD GOP ERROR MSG : ",Exception )
        elif i == "PCR[00] TPM 2.0 SHA256":
            SHA256_content = []
            if isAMDAMDPlatform :
                SHA256_file = re.compile(".*\d+_SHA256.txt")
                SHA256_file = list( filter( SHA256_file.match, allDir ) )
                if len(SHA256_file) > 0 :
                    with open(SHA256_file[0], encoding = "utf-16le") as f:
                            for line in f.readlines():
                                SHA256_content.append(line)
                    indexOfSHA = SHA256_content.index("TPM2_Startup: Return Code: 0x100\n")+1
                    sha256 = SHA256_content[indexOfSHA:indexOfSHA+2]
                    sha256[0] = sha256[0][8:-3]
                    sha256[1] = sha256[1][8:-2]
                    sha256 = sha256[0] + "\n" + sha256[1]
                    print(sha256)
                    outputFile[0].at[i, "Geted"] = sha256
                    continue
                else :
                    print("\nCan not get {any}[num]_SHA256.txt !")
            else :
                SHA256_file = re.compile("Custom Test Suite.log")
                SHA256_file = list( filter( SHA256_file.match, allDir ) )
                if len(SHA256_file) > 0 :
                    with open(SHA256_file[0]) as f:
                            for line in f.readlines():
                                SHA256_content.append(line)
                    sha256 = re.compile(".*PCR Index 00:.*")
                    sha256 = list( filter( sha256.match, SHA256_content ) )[0].split(":")[-1].strip()
                    outputFile[0].at[i, "Geted"] = sha256
                    print(sha256)
                    continue
                else :
                    print("\nCan not get Custom Test Suite.log !")
        elif i == "FUR":
            furV = ""
            if os.path.isfile(".\\HPFWUPDREC\\HpFirmwareUpdRec64.exe") :
                furV = information_parser.GetFileVersion(r".\\HPFWUPDREC\\HpFirmwareUpdRec64.exe")
            #print(furV)
            else :
                print("\nCan not get .\\HPFWUPDREC\\HpFirmwareUpdRec64.exe !")
            outputFile[0].at[i, "Geted"] = furV
            continue
        ###some should not get 
        elif i in {"Configurations", "1st", "2nd", "3rd", "Average"\
                    , "BIOS Initialization Duration(MS)", "Total Boot Duration(MS)"\
                    , "BiosInitTime(MS)", "DriverWakeTime(MS)"\
                    , "NOTE FOR THIS BIOS RELEASE", "TOOL REVISION"\
                    , "System POST TIME", "BIOS MODULE INFORMATION"\
                    , "BOOT TIME (ADK)", "S3 RESUME TIME"\
                    , "Known SI issues ready for retest with this release"\
                    , "Issue lists", "EC/SIO Functional changes" } :
            #print(i + " skip!")
            continue
        elif i == "ME Firmware":
            try :
                mef = bcu_content[bcu_content.index("ME Firmware Version\n")+1].strip()
                #print(mef)
                mef = "Corporate  v"+mef
                outputFile[0].at[i, "Geted"] = mef
                continue
            except :
                pass
        elif i == "Reference Code" :
            try :
                rc = bcu_content[bcu_content.index("Reference Code Revision\n")+1].strip()
                #print(rc)
                outputFile[0].at[i, "Geted"] = rc
                continue
            except :
                pass
        elif i == "Intel GOP EFI Driver" :
            try :
                igop = bcu_content[bcu_content.index("Video BIOS Version\n")+1].split("[")[-1].split("]")[0]
                outputFile[0].at[i, "Geted"] = igop
                continue
            except :
                pass
        elif i == "GbE Version" and not isAMDAMDPlatform :
            if binFile :
                with open(binFile, 'rb') as f:
                    f.seek(4106,0)
                    content1 = f.read(1)
                    content2 = f.read(1)
                if str(hex(ord(content2.decode()))).split("x")[-1] == "0" :
                    gbev = "00."+str(hex(ord(content1.decode()))).split("x")[-1]
                else :
                    gbev = str(hex(ord(content2.decode()))).split("x")[-1] \
                        +"."+str(hex(ord(content1.decode()))).split("x")[-1]
                print(gbev)
                outputFile[0].at[i, "Geted"] = gbev
            continue
        else : 
            pass
        print("Can not find : " + str(i) )
        outputFile[0].at[i, "Geted"] = "N/A"

outputFile[0]["ThisTimeUpdate"].fillna(value="N/A",inplace=True)
outputFile[0]["Geted"].fillna(value="N/A",inplace=True)

#######Check is same or not
for i in rRowInfoName:
    #print(i)
    if type(i) != str or i == "TOOL REVISION" or i == "NOTE FOR THIS BIOS RELEASE"\
        or i == "System POST TIME" or i == "BOOT TIME (ADK)" \
        or i == "S3 RESUME TIME" or i == "BIOS MODULE INFORMATION" \
        or i == "EC/SIO Functional changes" : # skip nan
        continue
    if str(outputFile[0].at[i, "ThisTimeUpdate"]) == "N/A" \
        and str(outputFile[0].at[i, "Geted"]) == "N/A" :
        outputFile[0].at[i, "Result"] = "Both N/A" 
    elif str(outputFile[0].at[i, "ThisTimeUpdate"]) == str(outputFile[0].at[i, "Geted"]) :
        outputFile[0].at[i, "Result"] = "V"  
    else :
        outputFile[0].at[i, "Result"] = "X"

#clean na
for i in range(len(rRowInfoName) ):
    if outputFile[0].iat[i,2] != outputFile[0].iat[i,2] :
        outputFile[0].iat[i,1] = " "
        outputFile[0].iat[i,0] = " "


#Save Excel
outputFile_PlatformHistory.to_excel(writer,sheet_name='PlatformHistory')
writer.close()
print("Comparison completed successfully!")
os.system("pause")
