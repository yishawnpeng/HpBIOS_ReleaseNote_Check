############################
#UTF-8
#Shawn.Peng@quantatw.com
#Can help you to double confirm info between release_note and BCU
#support relase_note is .xlsm not .docx(usually before G5)
############################
# Version List
# v1 base function
# v2 fix write sheet name error
# v3 Support G9/G9R in one release note.
# v4 Support AMI.
# v5 Try to get external link.
############################
# There are its base functions:
# 1.Get base information in BCU.txt witch show in BIOS R-N(exactly same name).
# 2.Try to get Check Sum from .bin.
# 3.Try to get Sprint(need local build).
# 4.Try to Get Agesa PI / SMU / PSP /… from amdz(if AMD).
# 5.Try to Get ME/RC/GbE… (if Intel).
# 6.Try to get SHA256.
# 7.Try to get FUR.exe version.
# (If 2~7  get fail, it still continue.)
# 8.And finally compare above 1~7 information then save the result.xlsm.
# Note : The BCU.txt / BIOS Release_Note.xlsm must exist and do not opening. The result.xlsm can not opening either(Maybe you run once).
#
# Share point
# xxx\CMIT_BIOS\Tools\compare_BCU_RN\compare_BCU_RN_V{number}.7z
# GitHub 
# https://github.com/yishawnpeng/HpBIOS_ReleaseNote_Check.git
#
############################
#pip3 install pandas re os openpyel pypiwin32 shutil

import pandas as pd     #excel
#import re               #regular expression    #imported from lib#
#import os               #dir                   #imported from lib#
import sys              #exit don't use os.exit
from win32com.client import * # GetFileVersion from exe

from lib import *

hihih

AMDPlatformDict = {"R24","R26","S25","S27","S29","T25","T26","T27"}
isAMDPlatform = None
AMIPlatformDict = {"U24"}
isAMIPlatform = None

if __name__=="__main__" :
    version = "5"
    #print("Version : "+version)
    arg=argparse_function(version)
    #Let user input platform and version
    goal_platform = input("\nInput Platform : ")
    goal_version = input("Input Version : ")
    if os.path.isdir(".\\"+str(goal_platform)+"_"+str(goal_version)):
        print("Go to folder : "+str(goal_platform)+"_"+str(goal_version))
        os.chdir(".\\"+str(goal_platform)+"_"+str(goal_version))
    else :
        print("Can not find "+str(goal_platform)+"_"+str(goal_version))
        os.system("pause")
        sys.exit()

    #Get FUR.exe info
    information_parser = Dispatch("Scripting.FileSystemObject")

    #Get Release Note File
    allDir = os.listdir( os.getcwd() ) #list
    excelName = re.compile("\w.*Release_Note_\d*\.xlsm|\w.*Release_Note.xlsm") 
    #{not~}{any}Release_Note_{number}.xlsm or {not~}{any}Release_Note.xlsm
    excelName = list( filter( excelName.match, allDir ) )
    """
    if not excelName : #empty
        #try AMI
        excelName = re.compile("\w.*Release_Note.xlsm") #{not~}{any}Release_Note.xlsm
        excelName = list( filter( excelName.match, allDir ) )
    """
    if not excelName : #empty
        print("Release_Note_excel is not exist or being opened!")
        print("Format should be {any}Release_Note_{number}.xlsm or {any}Release_Note.xlsm!")
        os.system("pause")
        sys.exit()
    else :
        logging.debug("Debug Mode")
        rName = excelName[0]
        platform = rName.split("_")[2]
        version = rName.split("_")[-1].split(".")[0]
        isR = True if rName.split("_")[0][-1] == "R" else False
        #print(version)
        print( "Choose release note: " + rName )

        #Check platfrom
        isAMIPlatform = True if platform in AMIPlatformDict else False
        isAMDPlatform = True if platform in AMDPlatformDict else False
        if isAMDPlatform :
            amdz_content = []
            getAMDZInfo(amdz_content)

        #Get item name and info of this time
        try :
            #Item name              :   usecols=[0]
            #Get from Release Note  :   usecols=[1]
            if isAMDPlatform :
                rRowInfoName = pd.read_excel( rName, sheet_name = "AMDPlatformHistory", usecols=[0] )
                rRowData = pd.read_excel( rName, sheet_name = "AMDPlatformHistory", usecols=[1] )
            elif platform in {"U21","U23"} :
                if isR :
                    rRowInfoName = pd.read_excel( rName, sheet_name = "IntelPlatformHistory_FY23", usecols=[0] )
                    rRowData = pd.read_excel( rName, sheet_name = "IntelPlatformHistory_FY23", usecols=[1] )
                else :
                    rRowInfoName = pd.read_excel( rName, sheet_name = "IntelPlatformHistory_FY22", usecols=[0] )
                    rRowData = pd.read_excel( rName, sheet_name = "IntelPlatformHistory_FY22", usecols=[1] )
            # include Intel AMI
            else : 
                rRowInfoName = pd.read_excel( rName, sheet_name = "IntelPlatformHistory", usecols=[0] )
                rRowData = pd.read_excel( rName, sheet_name = "IntelPlatformHistory", usecols=[1] )

            rRowInfoName = rRowInfoName[rRowInfoName.columns[0]].tolist()
            #Find Item Range
            startIndex = rRowInfoName.index("System BIOS Version")
            endIndex = rRowInfoName.index("Sprint Release Note")
        except Exception :
            #print(Exception)
            print("Get release note info! May be ceil(sheet) name error.")
            os.system("pause")
            sys.exit()

    #Create Excel
    try:
        writer = pd.ExcelWriter("result_RN.xlsx")
        outputFile = []
        outputFile_PlatformHistory = pd.DataFrame( index = rRowInfoName[startIndex:endIndex], \
                                                  columns = ["Release Note Info", "Reference Info", "Result"] )
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
    bcu_name = re.compile(".*BCU\.txt|.*bcu\.txt") # {any}BCU.txt or {any}bcu.txt
    bcu_name = list( filter( bcu_name.match, allDir ) )
    """
    if not bcu_name :   #Try .*bcu.txt
        bcu_name = re.compile(".*bcu\.txt") # {any}bcu.txt
        bcu_name = list( filter( bcu_name.match, allDir ) )
    """
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

        if isAMIPlatform :              # \w.*Release_Note.xlsm
            version = "".join(bcu_content[bcu_content.index("BIOS Revision\n")+1].strip()[1:].split("."))
        """
        if platform in {"U21","U23"} and isR :
            if not bcu_content[bcu_content.index("System Board ID\n")+1].strip() in {"8AC3","8AC6"} :
                print("Geted BCU SSID not in G9R(8AC3/8AC6)!")
                print("You should check!")
        elif platform in {"U21","U23"} :
            if bcu_content[bcu_content.index("System Board ID\n")+1].strip() in {"8AC3","8AC6"} :
                print("Geted BCU SSID in G9R(8AC3/8AC6)!")
                print("You should check!") 
        """

    #Find same in release note
    rRowInfoName = rRowInfoName[startIndex:endIndex]
    #print(bcu_content[ bcu_content.index(rRowInfoName[0]+"\n")+1])
    #print(bcu_content)
    #print(rRowInfoName)
    for i in rRowInfoName:
        if type(i) != str : # skip nan
            continue
        try :
            binFile = "" #checksun / GBE
            name = i +"\n"
            index = bcu_content.index(name)+1
            if i == "System BIOS Version" :
                temp = bcu_content[index].split()[2]
            elif i =="BIOS Build Version" :
                temp = bcu_content[index].strip()
            else :
                temp = bcu_content[index]
            outputFile[0].at[i, "Reference Info"] = temp
        except Exception :
            if i == "Build Date" and not isAMIPlatform :
                try:
                    bdate = bcu_content[bcu_content.index("System BIOS Version\n")+1].split()[-1]
                    outputFile[0].at[i, "Reference Info"] = bdate
                    continue
                except :
                    pass
            elif i == "CHECKSUM" and not isAMIPlatform :
                #AMD
                if isAMDPlatform and os.path.isfile(".\\Global\\BIOS\\" + platform + "_" + version + ".bin") :
                    binFile = ".\\Global\\BIOS\\" + platform + "_" + version + ".bin"
                #Intel
                elif not isAMDPlatform and os.path.isfile(".\\Global\\BIOS\\" + platform + "_" + version + "_32.bin") :
                    binFile = ".\\Global\\BIOS\\" + platform + "_" + version + "_32.bin"

                if binFile :
                    with open(binFile, 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                        f.close()
                    #x need lower, other need upper
                    binary_sum = "0x"+binary_sum.split("x")[-1].upper()
                    outputFile[0].at[i, "Reference Info"] = binary_sum
                else :
                    print("\nCan note find biniary : "+platform + "_" + version + ".bin")
                    print("Format : {Platform}_{version}.bin (Fist letter upper) ")
                    print("NOTE:\nAMD:.bin ; Intel:_32.bin")
                continue
            elif i == "Sprint" :
                print("\nSprint info in local build BCU!")
            elif i == "EC/SIO F/W" and not isAMIPlatform :
                try:
                    sio = bcu_content[bcu_content.index("Super I/O Firmware Version\n")+1].split()[-1]
                    outputFile[0].at[i, "Reference Info"] = sio
                    continue
                except :
                    pass
            elif isTypecPD(i) :
                try :
                    firstPD = bcu_content[bcu_content.index("USB Type-C Controller(s) Firmware Version:\n")+1].split()[-1]
                    secondPD = bcu_content[bcu_content.index("USB Type-C Controller(s) Firmware Version:\n")+2].split()[-1]
                    numdot=re.compile("[0-9]+")
                    if not numdot.match(firstPD) :
                        outputFile[0].at[i, "Reference Info"] = "N\A"
                    #check secondPD is exit or not
                    elif numdot.match(firstPD) and not numdot.match(secondPD) :
                        outputFile[0].at[i, "Reference Info"] = firstPD
                        continue
                    else :
                        outputFile[0].at[i, "Reference Info"] = firstPD + "\n" + secondPD
                        continue
                except :
                    pass
            elif i == "PCR[00] TPM 2.0 SHA256":
                SHA256_content = []
                if isAMDPlatform :
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
                        outputFile[0].at[i, "Reference Info"] = sha256
                        continue
                    else :
                        print("\nCan not find {any}[num]_SHA256.txt !")
                else :
                    SHA256_file = re.compile("Custom Test Suite.log")
                    SHA256_file = list( filter( SHA256_file.match, allDir ) )
                    if len(SHA256_file) > 0 :
                        with open(SHA256_file[0]) as f:
                                for line in f.readlines():
                                    SHA256_content.append(line)
                        sha256 = re.compile(".*PCR Index 00:.*")
                        sha256 = list( filter( sha256.match, SHA256_content ) )[0].split(":")[-1].strip()
                        outputFile[0].at[i, "Reference Info"] = sha256
                        continue
                    else :
                        print("\nCan not find Custom Test Suite.log !")
            elif i == "FUR" and not isAMIPlatform:
                furV = ""
                if os.path.isfile(".\\HPFWUPDREC\\HpFirmwareUpdRec64.exe") :
                    furV = information_parser.GetFileVersion(r".\\HPFWUPDREC\\HpFirmwareUpdRec64.exe")
                else :
                    print("\nCan not find .\\HPFWUPDREC\\HpFirmwareUpdRec64.exe !")
                outputFile[0].at[i, "Reference Info"] = furV
                continue
            elif i == "SVN ver. Core" :
                el_name = re.compile(".*External_Link\.txt|.*EL\.txt") # {any}External_Link.txt or {any}EL.txt
                el_name = list( filter( el_name.match, allDir ) )
                if not el_name :
                    print("\nCan not find {any}External_Link.txt")
                    print("Format : {any}External_Link.txt or {any}EL.txt ")
                else :
                    el_content=[]
                    with open(el_name[0]) as f:
                        for line in f.readlines():
                            el_content.append(line)
                    if ("R" in platform) or ("S" in platform) :     #G5/G6
                        svn_core = re.compile(".*HpCore\n")
                    else : # ("" in platform)                       #G8 git
                        svn_core = re.compile(".*HpCorePvtBins\n")
                    svn_core = list( filter( svn_core.match, el_content ) )[0].split()[1]
                    outputFile[0].at[i, "Reference Info"] = svn_core
                    continue
            elif i == "SVN ver. Chipset & PE" :
                if el_name :
                    if isAMDPlatform :                              #AMD
                        svn_pi = re.compile(".*AMD\n")
                    else :                                          #Intel
                        svn_pi = re.compile(".*HpIntelChipsetPkg\n")
                    svn_pi = list( filter( svn_pi.match, el_content ) )[0].split()[1]
                    outputFile[0].at[i, "Reference Info"] = svn_pi
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
                continue
            ##########AMD start
            elif i == "AMD Agesa PI" :
                if amdz_content :
                    agesaPI = re.compile("AGESA:.*")
                    agesaPI = list( filter( agesaPI.match, amdz_content ) )[0].split()[-1] #agesaPI = agesaPI[0].split()[-1]
                    outputFile[0].at[i, "Reference Info"] = agesaPI
                    continue
                else :
                    print("\nCan not find PI because not have amdz!")
            elif i == "PSP FW" :
                if amdz_content :
                    PSPandSMU = re.compile("SMU:.*")
                    PSPandSMU = list( filter( PSPandSMU.match, amdz_content ) )
                    PSPandSMU = PSPandSMU[0].split()
                    pspfw = PSPandSMU[2]
                    if "(" in pspfw  :
                        pspfw = pspfw.split("(")[1][:-1]
                    else :
                        pspfw = pspfw[2:]
                    realPSPFW = ""
                    for j in range(0,len(pspfw),2) :
                        if pspfw[j] == "0" and pspfw[j+1] == "0" :
                            realPSPFW = realPSPFW + "0."
                        elif pspfw[j] == "0" and pspfw[j+1] != "0" :
                            realPSPFW = realPSPFW + pspfw[j+1] + "."
                        else : 
                            realPSPFW = realPSPFW + pspfw[j] + pspfw[j+1] + "."
                    realPSPFW = realPSPFW[:-1]
                    outputFile[0].at[i, "Reference Info"] = realPSPFW
                    continue
                else :
                    print("\nCan not find PSP FW because not have amdz!")
            elif i == "SMU FW" : #maybe dec to hex
                if amdz_content :
                    smufw = PSPandSMU[1].split("(")[0]
                    outputFile[0].at[i, "Reference Info"] = smufw
                    continue
                else :
                    print("\nCan not find SMU FW because not have amdz!")
            elif i == "AMD Legacy VBIOS" :
                if amdz_content :
                    vBIOS = re.compile("VBIOS Info.*")
                    vBIOS = list( filter( vBIOS.match, amdz_content ) )[0].split()[3][0:-1]
                    outputFile[0].at[i, "Reference Info"] = vBIOS
                    continue
                else :
                    print("\nCan not find SMU FW because not have amdz!")
            elif i == "AMD GOP EFI Driver" :
                try :
                    gOP = re.compile("Rev.*")
                    gOP = list( filter( gOP.match, bcu_content[bcu_content.index("Video BIOS Version\n")+1].split() ) )
                    gOP = gOP[0][4:-4]
                    outputFile[0].at[i, "Reference Info"] = gOP
                    continue
                except Exception:
                    print("AMD GOP ERROR MSG : ",Exception )
            ##########AMD end
            ##########Intel start
            elif i == "ME Firmware" and not isAMIPlatform:
                try :
                    mef = bcu_content[bcu_content.index("ME Firmware Version\n")+1].strip()
                    mef = "Corporate  v"+mef
                    outputFile[0].at[i, "Reference Info"] = mef
                    continue
                except :
                    pass
            elif i == "Reference Code" and not isAMIPlatform :
                try :
                    rc = bcu_content[bcu_content.index("Reference Code Revision\n")+1].strip()
                    outputFile[0].at[i, "Reference Info"] = rc
                    continue
                except :
                    pass
            elif i == "Intel GOP EFI Driver" and not isAMIPlatform :
                try :
                    igop = bcu_content[bcu_content.index("Video BIOS Version\n")+1].split("[")[-1].split("]")[0]
                    outputFile[0].at[i, "Reference Info"] = igop
                    continue
                except :
                    pass
            elif i == "GbE Version" and not isAMIPlatform :
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
                    outputFile[0].at[i, "Reference Info"] = gbev
                continue
            elif i == "Processor Microcode Patches" and not isAMIPlatform :
                try :
                    pm = bcu_content[bcu_content.index("Processor 1 MicroCode Revision\n")+1].strip()
                    outputFile[0].at[i, "Reference Info"] = pm
                    continue
                except :
                    pass
            ##########Intel end
            ##########AMI start
            elif isAMIPlatform :
                try :
                    #print("AMI!! "+i)
                    if i == "System BIOS Version" :
                        biosversion = bcu_content[bcu_content.index("BIOS Revision\n")+1].strip()[1:]
                        outputFile[0].at[i, "Reference Info"] = biosversion
                        continue
                    elif i == "Build Date" :
                        try :
                            bdate = bcu_content[bcu_content.index("BIOS Date (ReadOnly)\n")+1].strip()
                        except :
                            try :
                                bdate = bcu_content[bcu_content.index("BIOS Date \n")+1].strip()
                            except :
                                bdate = bcu_content[bcu_content.index("BIOS Date\n")+1].strip()
                        outputFile[0].at[i, "Reference Info"] = bdate
                        continue
                    elif i == "CHECKSUM" :
                        AMIbinFile = ""
                        if os.path.isfile(".\\Global\\BIOS\\" + platform + "_" + version + "_16.bin") :
                            AMIbinFile = ".\\Global\\BIOS\\" + platform + "_" + version + "_16.bin"
                        if AMIbinFile :
                            with open(AMIbinFile, 'rb') as f:
                                content = f.read()
                                binary_sum = sum(bytearray(content))
                                binary_sum = hex(binary_sum & 0xFFFFFFFF)
                                f.close()
                            binary_sum = binary_sum.split("x")[-1].upper()
                            outputFile[0].at[i, "Reference Info"] = binary_sum
                        else :
                            print("\nCan note find biniary : "+platform + "_" + version + "_16.bin")
                            print("AMI bin Format : {Platform}_{version}_16.bin (Fist letter upper).")
                        continue
                    elif i == "FUR" :
                        if os.path.isfile(platform + "_" + version + ".exe") :
                            furV = information_parser.GetFileVersion(platform + "_" + version + ".exe")
                        else :
                            print("\nCan not find .\\U24_[version].exe !")
                        outputFile[0].at[i, "Reference Info"] = furV
                        continue
                    elif i == "ME Firmware" :
                        try :
                            me_name = re.compile(r"ME_[0-9\.]+\.bin") 
                            ori_dir = os.getcwd()
                            os.chdir(".\\METools\\FWUpdate\\MEFW")
                            me_dir = os.listdir( os.getcwd() ) #list
                            me_name = list( filter( me_name.match, me_dir ) )
                            os.chdir(ori_dir)
                        except :
                            print("Get AMI ME folder fail !")
                        if len(me_name) > 0 :
                            mef = me_name[0][3:-4]+"_Consumer"  #ME_[0-9\.]+\.bin
                        else :
                            print("\nCan not find .\\METools\\FWUpdate\\MEFW\\ME_{version}.bin !")
                        outputFile[0].at[i, "Reference Info"] = mef
                        continue
                    elif i == "GbE Version" :
                        """ # need to discuss
                        if os.path.isfile(".\\Global\\BIOS\\" + platform + "_" + version + "_16.bin") :
                            with open(".\\Global\\BIOS\\" + platform + "_" + version + "_16.bin", 'rb') as f:
                                f.seek(8202,0)
                                content1 = f.read(1)
                                content2 = f.read(1)
                            gbev = str(hex(ord(content2.decode()))).split("x")[-1] \
                                    +"."+str(hex(ord(content1.decode()))).split("x")[-1]
                            outputFile[0].at[i, "Reference Info"] = gbev
                        """
                        continue
                    else :
                        pass
                except :
                    pass
            ##########AMI end
            else : 
                pass
            print("Can not find : " + str(i) )
            outputFile[0].at[i, "Reference Info"] = "N/A"

    outputFile[0]["Release Note Info"].fillna(value="N/A",inplace=True)
    outputFile[0]["Reference Info"].fillna(value="N/A",inplace=True)

    #######Check is same or not
    for i in rRowInfoName:
        if type(i) != str or i == "TOOL REVISION" or i == "NOTE FOR THIS BIOS RELEASE"\
            or i == "System POST TIME" or i == "BOOT TIME (ADK)" \
            or i == "S3 RESUME TIME" or i == "BIOS MODULE INFORMATION" \
            or i == "EC/SIO Functional changes" : # skip nan
            continue
        if str(outputFile[0].at[i, "Release Note Info"]) == "N/A" \
            and str(outputFile[0].at[i, "Reference Info"]) == "N/A" :
            outputFile[0].at[i, "Result"] = "Both N/A" 
        elif str(outputFile[0].at[i, "Release Note Info"]) == str(outputFile[0].at[i, "Reference Info"]) :
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
    print("\nComparison completed successfully!\n")
    os.system("pause")
