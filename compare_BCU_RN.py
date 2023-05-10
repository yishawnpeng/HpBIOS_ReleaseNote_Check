############################
#UTF-8
#Shawn.Peng@quantatw.com
#Can help you to double confirm info between release_note and BCU
############################
# There are its base functions:
# 1.Get base information in BCU.txt witch show in BIOS R-N(exactly same name).
# 2.Try to get Check Sum from .bin.
# 3.Try to get Sprint(need local build).
# 4.Try to Get Agesa PI / SMU / PSP /… from amdz(if AMD).
# 5.Try to Get ME/RC/GbE… (if Intel).
# 6.Try to get Exteranl Link version.
# 7.Try to get SHA256.
# 8.Try to get FUR.exe version.
# (If 4~6 get fail, it still continue.)
# 9.And finally compare above 1~8 information then save the result.xlsm.
# Note : The BCU.txt / BIOS Release_Note.xlsm must exist and do not opening. The result.xlsm can not opening either(Maybe you run once).
#
# Share point
# xxx\CMIT_BIOS\Tools\compare_BCU_RN\compare_BCU_RN_V{number}.7z
# GitHub 
# https://github.com/yishawnpeng/HpBIOS_ReleaseNote_Check.git
#
############################
import pandas as pd     #excel
import sys              #exit don't use os.exit
import shutil           #for copy file (os.rename will remove file)
import docx
from win32com.client import * # GetFileVersion from exe
from lib import *

version = "7"
#support G3/G4 (release note docx)
arg=argparse_function(version)

AMDPlatformDict = {"R24","R26","S25","S27","S29","T25","T26","T27"}
AMDG4PlatformDict = {"Q26","Q27"}
isAMDPlatform = None
isAMDG4Platform = None
AMIPlatformDict = {"U24"}
isAMIPlatform = None
isG4Platform = None
###print("ME is compare form BCU not .bin")

#=========Let user input platform and version
goal_platform = input("\nInput Platform : ")
goal_version = input("Input Version : ")
release_dir = ".\\"+str(goal_platform)+"_"+str(goal_version)
fatherDir = os.getcwd() # father / outside
allDir = os.listdir( fatherDir ) #list
#====Check G4
if "Q" in goal_platform or "P" in goal_platform :
    if goal_platform in AMDG4PlatformDict :
        isAMDG4Platform = True
    else :
        isG4Platform = True
#====Create folder
if os.path.isdir(release_dir):
    os.chdir(release_dir)
    release_dir = os.getcwd()
    os.chdir(fatherDir)
    new_dir = ".\\"+str(goal_platform)+"_"+str(goal_version)+"_checked"
    if not os.path.exists(new_dir) :
        print("Create folder:"+new_dir)
        os.makedirs(new_dir)
        os.chdir(new_dir)
        new_dir = os.getcwd()
        os.chdir(fatherDir)
    else :
        print("Folder already exist:"+new_dir)
        os.chdir(new_dir)
        new_dir = os.getcwd()
        os.chdir(fatherDir)
else :
    print("Can not find "+release_dir)
    os.system("pause")
    sys.exit()
#====Move file (from outside)
#BCU
bcu_name = re.compile(".*BCU\.txt|.*bcu\.txt") # {any}BCU.txt or {any}bcu.txt
bcu_name = list( filter( bcu_name.match, allDir ) )
if not bcu_name :   #empty
    print("You don't have BCU file or being opened!\nFormat should be {any}BCU.txt or {any}bcu.txt !")
    os.system("pause")
    sys.exit()
else :
    print("Choose BCU :" + bcu_name[0])
    if not os.path.isfile(new_dir+"\\"+bcu_name[0]) :
        os.rename(fatherDir+"\\"+bcu_name[0],new_dir+"\\"+bcu_name[0])
#AMDZ
amdz_name = re.compile("amdz.*\.txt")
amdz_name = list( filter( amdz_name.match, allDir ) )
if not amdz_name : # empty
    print("You don't have amdz file !\nFormat : amdz{any}.txt")
else :
    print("Choose amdz :" + amdz_name[0])
    if not os.path.isfile(new_dir+"\\"+amdz_name[0]) :
        os.rename(fatherDir+"\\"+amdz_name[0],new_dir+"\\"+amdz_name[0])
#External Link
el_name = re.compile(".*External_Link\.txt|.*EL\.txt") # {any}External_Link.txt or {any}EL.txt
el_name = list( filter( el_name.match, allDir ) )
if not el_name :
    print("Can not find {any}External_Link.txt!\nFormat : {any}External_Link.txt or {any}EL.txt")
else :
    print("Choose External Link :" + el_name[0])
    if not os.path.isfile(new_dir+"\\"+el_name[0]) :
        os.rename(fatherDir+"\\"+el_name[0],new_dir+"\\"+el_name[0])
#====Copy file (from release file)
os.chdir(release_dir)
release_all_dir = os.listdir( os.getcwd() )
#Release Note
excelName = re.compile("\w.*Release_Note_\d*\.xlsm|\w.*Release_Note.xlsm") 
if isAMDG4Platform or isG4Platform :
    excelName = re.compile("\w.*release note.docx|\w.*_Release_Notes.docx") 
excelName = list( filter( excelName.match, release_all_dir ) )
if not excelName :
    if isAMDG4Platform or isG4Platform :
        print("Can not find Release Note!\nFormat :{any}release note.docx or {any}_Release_Notes.docx")
    else :
        print("Can not find Release Note!\nFormat :{any}Release_Note_{number}.xlsm or {any}Release_Note.xlsm")
    os.system("pause")
    sys.exit()
else :
    print("Choose Release Note :" + excelName[0])
    shutil.copy(release_dir+"\\"+excelName[0], new_dir+"\\"+excelName[0])
#SHA256
SHA256_file = re.compile(".*\d+_SHA256.txt")
SHA256_file = list( filter( SHA256_file.match, release_all_dir ) )
if not SHA256_file :
    print("Can not find {any}[number]_SHA256.txt !")
else :
    print("Choose SHA256 :" + SHA256_file[0])
    shutil.copy(release_dir+"\\"+SHA256_file[0], new_dir+"\\"+SHA256_file[0])
#FUR (from HPFWUPDREC)
information_parser = Dispatch("Scripting.FileSystemObject")
furP = ".\\HPFWUPDREC\\HpFirmwareUpdRec64.exe"
if os.path.isfile(furP) :
    shutil.copy(release_dir+"\\"+furP, new_dir+"\\"+"HpFirmwareUpdRec64.exe")
    furP = "HpFirmwareUpdRec64.exe"
    #furV = information_parser.GetFileVersion(r".\\HPFWUPDREC\\HpFirmwareUpdRec64.exe")
else :
    ##AMI start
    furP=""
    ami_furP = re.compile("U24_\d+.exe")
    ami_furP = list( filter( ami_furP.match, release_all_dir ) )
    if os.path.isfile(ami_furP[0]) :
        shutil.copy(release_dir+"\\"+ami_furP[0], new_dir+"\\"+ami_furP[0])
        ami_furP = ami_furP[0]
        #furV = information_parser.GetFileVersion(platform + "_" + version + ".exe")
    else :
        print("Can not find .\\U24_[number].exe !")
    ##AMI end
    print("Can not find .\\HPFWUPDREC\\HpFirmwareUpdRec64.exe !")
#bin (for checksum from Global/BIOS)
if os.path.isfile(".\\Global\\BIOS\\"+goal_platform+"_"+goal_version+".bin") :
    #AMD
    shutil.copy(release_dir+"\\Global\\BIOS\\"+goal_platform+"_"+goal_version+".bin"\
              , new_dir+"\\"+goal_platform+"_"+goal_version+".bin")
    binFile = goal_platform+"_"+goal_version+".bin"
elif os.path.isfile(".\\Global\\BIOS\\"+goal_platform+"_"+goal_version+"_16.bin") :
    #AMI
    shutil.copy(release_dir+"\\Global\\BIOS\\"+goal_platform+"_"+goal_version+"_16.bin"\
              , new_dir+"\\"+goal_platform+"_"+goal_version+"_16.bin")
    binFile = goal_platform+"_"+goal_version+".bin"
elif os.path.isfile(".\\Global\\BIOS\\"+goal_platform+"_"+goal_version+"_32.bin") :
    #Intel
    shutil.copy(release_dir+"\\Global\\BIOS\\"+goal_platform+"_"+goal_version+"_32.bin"\
              , new_dir+"\\"+goal_platform+"_"+goal_version+"_32.bin")
    binFile = goal_platform+"_"+goal_version+".bin"
else :
    print("Can not find {platform}_{version}_[|16|32].bin !")
#=========Compare info
#Go to new folder
os.chdir(new_dir)
#Get BCU info
bcu_content=[]
with open(bcu_name[0]) as f:
    for line in f.readlines():
        bcu_content.append(line)
#Get Release Note info
if isAMDG4Platform or isG4Platform :
    print("a1")
    rName = excelName[0]
    if isAMDG4Platform :
        platform = rName.split("_")[1]
    else :
        platform = rName.split("_")[2]
    if isAMDG4Platform :
        version = rName.split("_")[2]
        version = version.replace(".","")
    else :
        version = rName.split("_")[3]
        version = version.split(" ")[0]
    if goal_platform != str(platform) or goal_version != str(version) :
        print("\nYour INPUT plateform_version is different with geted release note plateform_version!\nYou might ckeck!\n")
    try :
        """ G4
        file  = docx.Document("Scotty_Q26_02.22.00_0001_Release_Notes.docx")
        #intel Sax_PV_Q11_022201 release note.docx
        #print("len",len(file.paragraphs))
        tables=file.tables
        table = tables[1]
        for i in range(0,len(table.rows)) :
            result = table.cell(i,0).text
            r2 = table.cell(i,1).text
            print(result)
            print(r2)
        """
        print("a2")
        rRowInfoName = docx.Document(rName)
        table=rRowInfoName.tables[1]
        rRowInfoName=[]
        rRowData=[]
        print("a3")
        for i in range(0,len(table.rows)) :
            print("a4")
            rRowInfoName.append(table.cell(i,0).text)
            rRowData.append(table.cell(i,1).text)

        print("a5")
        #Find Item Range
        startIndex = rRowInfoName.index("System BIOS")
        endIndex = rRowInfoName.index("CHID")
    except Exception :
        print("a6")
        #print(Exception)
        print("Get release note info! May be ceil(sheet) name error.")
        os.system("pause")
        sys.exit()
else :
    print("b")
    rName = excelName[0]
    platform = rName.split("_")[2]
    version = rName.split("_")[-1].split(".")[0]
    isR = True if rName.split("_")[0][-1] == "R" else False
    isAMIPlatform = True if goal_platform in AMIPlatformDict else False
    if isAMIPlatform :              # \w.*Release_Note.xlsm
        version = "".join(bcu_content[bcu_content.index("BIOS Revision\n")+1].strip()[1:].split("."))
    isAMDPlatform = True if goal_platform in AMDPlatformDict else False
    if goal_platform != str(platform) or goal_version != str(version) :
        print("\nYour INPUT plateform_version is different with geted release note plateform_version!\nYou might ckeck!\n")
    ##Get item name and info of this time
    try :
        print("b1")
        #Item name              :   usecols=[0]
        #Get from Release Note  :   usecols=[1]
        if isAMDPlatform :
            print("b2")
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
            print("b3")
            rRowInfoName = pd.read_excel( rName, sheet_name = "IntelPlatformHistory", usecols=[0] )
            rRowData = pd.read_excel( rName, sheet_name = "IntelPlatformHistory", usecols=[1] )

        print("b4")
        rRowInfoName = rRowInfoName[rRowInfoName.columns[0]].tolist()
        #Find Item Range
        startIndex = rRowInfoName.index("System BIOS Version")
        endIndex = rRowInfoName.index("Sprint Release Note")
        print("b5")
    except Exception :
        print("b6")
        #print(Exception)
        print("Get release note info! May be ceil(sheet) name error.")
        os.system("pause")
        sys.exit()
#Get AMDZ info
if amdz_name :
    amdz_content=[]
    with open(amdz_name[0]) as f:
        for line in f.readlines():
            amdz_content.append(line)
#Get SHA256 info
if SHA256_file :
    SHA256_content=[]
    with open(SHA256_file[0], encoding = "utf-16le") as f:
        for line in f.readlines():
            SHA256_content.append(line)
#Get External Link
if el_name:
    el_content = []
    with open(el_name[0]) as f:
        for line in f.readlines():
            el_content.append(line)
#Create resault.xml
try:
    writer = pd.ExcelWriter("result_RN.xlsx")
    outputFile = []
    outputFile_PlatformHistory = pd.DataFrame( index = rRowInfoName[startIndex:endIndex], \
                                                columns = ["Release Note Info", "Reference Info", "Result"] )
    outputFile.append(outputFile_PlatformHistory) #Sheet No.1
    if isG4Platform or isAMDG4Platform :
        outputFile[0].iloc[:, 0] = rRowData[startIndex:endIndex]
    else :
        outputFile[0].iloc[:, 0] = rRowData[rRowData.columns[0]].tolist()[startIndex:endIndex]
except Exception:
    #print(Exception)
    print("Creat excel fail or result_RN.xlsx being opened!")
    os.system("pause")
    sys.exit()
#common
rRowInfoName = rRowInfoName[startIndex:endIndex]
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
            if binFile :
                with open(binFile, 'rb') as f:
                    content = f.read()
                    binary_sum = sum(bytearray(content))
                    binary_sum = hex(binary_sum & 0xFFFFFFFF)
                    f.close()
                #x need lower, other need upper
                binary_sum = "0x"+binary_sum.split("x")[-1].upper()
                outputFile[0].at[i, "Reference Info"] = binary_sum
            continue
        elif i == "Sprint" :
            print("Sprint info in local build BCU!")
        elif (i == "EC/SIO F/W" or i == "SIO FW") and not isAMIPlatform :
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
        elif i == "PCR[00] TPM 2.0 SHA256" or i == "PCR 0" :
            if SHA256_content :
                indexOfSHA = SHA256_content.index("TPM2_Startup: Return Code: 0x100\n")+1
                sha256 = SHA256_content[indexOfSHA:indexOfSHA+2]
                sha256[0] = sha256[0][8:-3]
                sha256[1] = sha256[1][8:-2]
                sha256 = sha256[0] + "\n" + sha256[1]
                outputFile[0].at[i, "Reference Info"] = sha256
                continue
        elif (i == "FUR" or i == "HPBIOSUPDREC") and not isAMIPlatform:
            furV = ""
            if furP :
                furV = information_parser.GetFileVersion(furP)
            outputFile[0].at[i, "Reference Info"] = furV
            continue
        elif i == "SVN ver. Core" :
            if el_name :
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
        elif i == "AMD Agesa PI" or i == "AMD Agesa code" :
            if amdz_name :
                agesaPI = re.compile("AGESA:.*")
                agesaPI = list( filter( agesaPI.match, amdz_content ) )[0].split()[-1] #agesaPI = agesaPI[0].split()[-1]
                outputFile[0].at[i, "Reference Info"] = agesaPI
                continue
        elif i == "PSP FW" :
            if amdz_name :
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
        elif i == "SMU FW" : #maybe dec to hex
            if amdz_name :
                smufw = PSPandSMU[1].split("(")[0]
                outputFile[0].at[i, "Reference Info"] = smufw
                continue
        elif i == "AMD Legacy VBIOS" or i == "AMD VBIOS" :
            if amdz_name :
                vBIOS = re.compile("VBIOS Info.*")
                vBIOS = list( filter( vBIOS.match, amdz_content ) )[0].split()[3][0:-1]
                outputFile[0].at[i, "Reference Info"] = vBIOS
                continue
        elif i == "AMD GOP EFI Driver" or i == "AMD GOP" :
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
        elif i == "GbE Version" and not isAMIPlatform and not isAMDPlatform :
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
                outputFile[0].at[i, "Reference Info"] = "0x"+pm
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
                    if binFile :
                        with open(binFile, 'rb') as f:
                            content = f.read()
                            binary_sum = sum(bytearray(content))
                            binary_sum = hex(binary_sum & 0xFFFFFFFF)
                            f.close()
                        binary_sum = binary_sum.split("x")[-1].upper()
                        outputFile[0].at[i, "Reference Info"] = binary_sum
                    continue
                elif i == "FUR" :
                    if ami_furP :
                        furV = information_parser.GetFileVersion(ami_furP)
                    outputFile[0].at[i, "Reference Info"] = furV
                    continue
                elif i == "ME Firmware" :
                    try :
                        me_name = re.compile(r"ME_[0-9\.]+\.bin") 
                        os.chdir(release_dir+"\\METools\\FWUpdate\\MEFW")
                        me_dir = os.listdir( os.getcwd() ) #list
                        me_name = list( filter( me_name.match, me_dir ) )
                        os.chdir(new_dir)
                    except :
                        print("Get AMI ME folder fail !")
                    if len(me_name) > 0 :
                        mef = me_name[0][3:-4]+"_Consumer"  #ME_[0-9\.]+\.bin
                    else :
                        print("\nCan not find .\\METools\\FWUpdate\\MEFW\\ME_{version}.bin !")
                    outputFile[0].at[i, "Reference Info"] = mef
                    continue
                elif i == "GbE Version" :
                    continue
                else :
                    pass
            except :
                pass
        ##########AMI end
        ##########G4 end
        elif i == "System BIOS" :
            try:
                bversion = bcu_content[bcu_content.index("System BIOS Version\n")+1].split()[2]
                outputFile[0].at[i, "Reference Info"] = "Ver " + bversion
                continue
            except :
                pass
            
        ##########G4 end
        else : 
            pass
        print("Can not find : " + str(i) )
        outputFile[0].at[i, "Reference Info"] = "N/A"

outputFile[0]["Release Note Info"].fillna(value="N/A",inplace=True)
outputFile[0]["Reference Info"].fillna(value="N/A",inplace=True)
#=========Resault
#compare
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
#====Safe file
outputFile_PlatformHistory.to_excel(writer,sheet_name='PlatformHistory')
writer.close()
print("\nComparison completed successfully!\n")
os.system("pause")