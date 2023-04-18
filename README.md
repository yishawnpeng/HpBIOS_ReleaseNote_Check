# HpBIOS_ReleaseNote_Check
compare_BCU_RN.py is python script that can help you to double confirm info between release_note and BCU(and other unnecessary file)

Support Intel G5/G6/G8/G9/G9R, AMD G5/G6/G8 and AMI consumer platforms.

### MAIN 
![example workflow](https://github.com/yishawnpeng/HpBIOS_ReleaseNote_Check/actions/workflows/Python_build.yml/badge.svg)

### RELEASE
![example workflow](https://github.com/yishawnpeng/HpBIOS_ReleaseNote_Check/actions/workflows/Release-build.yml/badge.svg)

## Main Function
Get base information in BCU.txt witch show in BIOS Release Note(exactly same name).
Get Check Sum from .bin.
Get SHA256.
Get FUR.exe version.

## Other Feature
Get following Info fail can also work.
* Try to get Sprint(need local build).
* Try to Get Agesa PI / SMU / PSP /… from amdz(if AMD).
* Try to Get ME/RC/GbE… (if Intel).
* Try to get External Link.

## Installation
1. Clone the repository: ```git clone https://github.com/yishawnpeng/HpBIOS_ReleaseNote_Check.git```
2. Install Python 3.x or later: https://www.python.org/downloads/
3. Install required libraries: ```pip install -r requirements.txt```

## How to use
### Before v5 
1. Put .exe <font color=Red>in</font> your release folder.
2. Put {any}BCU.txt or {any}bcu.txt in your release folder.
3. You can put {any}_SHA256.txt or Custom Test Suite.log / {any}_External_Link.txt or {any}_EL.txt / amdz{any}.txt in too(This step skip is FINE).
4. Run .exe (double click/ power shell / cmd ).
5. Check “result_RN.xlsx” file.

### After v5
1. Put .exe <font color=Red>out</font> of your release folder.
2. Put {any}BCU.txt or {any}bcu.txt out of release folder.
3. You can put {any}_External_Link.txt or {any}_EL.txt / amdz{any}.txt in too(This step skip is FINE).
4. Run .exe (double click/ power shell / cmd ).
5. Check “{platform}_checked/result_RN.xlsx” file.


## Contributing
If you would like to contribute to this project, please follow these steps:
 1. Fork the repository
 2. Create a new branch for your feature: ```git checkout -b feature-name```
 3. Make changes and commit them: ```git commit -am 'Add some feature'```
 4. Push to the branch: ```git push origin feature-name```
 5. Submit a pull request
