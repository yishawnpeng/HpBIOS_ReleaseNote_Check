############################
#UTF-8
#Shawn.Peng@quantatw.com
#be imported to CheckReleaseNote.py
############################
import re
import os
import argparse
import logging

def argparse_function(ver):
    parser = argparse.ArgumentParser(prog='compare_BCU_RN.py', description='Tutorial')
    parser.add_argument("-d", "--debug", help='Show debug message.', action="store_true")
    parser.add_argument("-v", "--version", action="version", version=ver)
    args = parser.parse_args()
    if args.debug:
        Debug_Format = "%(levelname)s, %(funcName)s: %(message)s"
        logging.basicConfig(level=logging.DEBUG, format=Debug_Format)  #Debug use flag
        print("Enable debug mode.")
    return ver

def isTypecPD(string) :
    amdN = re.compile("Cypress PD FW.*")
    intelN = re.compile("USB TYPE-C FW.*")
    if amdN.match(string) or intelN.match(string) :
        return True
    return False
