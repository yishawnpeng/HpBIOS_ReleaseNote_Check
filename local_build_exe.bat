@echo off
if exist compare_BCU_RN.spec (
    echo Found spec file, using it to build...
    pyinstaller compare_BCU_RN.spec
) else (
    echo No spec file, using .py to build...
    pyinstaller -F compare_BCU_RN.py
)

if %errorlevel% == 0 (
    echo Build Pass : please check \dist\compare_BCU_RN.exe

    if exist compare_BCU_RN.exe (
        echo Found existing compare_BCU_RN.exe in current folder, renaming it...
        ren compare_BCU_RN.exe compare_BCU_RN_o.exe
    )

    copy /Y dist\compare_BCU_RN.exe .\
) else (
    echo Fail %errorlevel%
)
pause