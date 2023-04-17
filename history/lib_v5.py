import re
import os

allDir = os.listdir( os.getcwd() ) #list
excelName = re.compile("\w.*Release_Note_\d*\.xlsm") #{not~}{any}Release_Note_{number}.xlsm
excelName = list( filter( excelName.match, allDir ) )

def getAMDZInfo(amdzContent) :
    amdz_name = re.compile("amdz.*\.txt")
    amdz_name = list( filter( amdz_name.match, allDir ) )
    if not amdz_name : # empty
        print("You don't have amdz file !")
        print("Format : amdz{any}.txt")
        return False
    else :
        with open(amdz_name[0]) as f:
            for line in f.readlines():
                amdzContent.append(line)
        return True
    
def isTypecPD(string) :
    amdN = re.compile("Cypress PD FW.*")
    intelN = re.compile("USB TYPE-C FW.*")
    if amdN.match(string) or intelN.match(string) :
        return True
    return False
