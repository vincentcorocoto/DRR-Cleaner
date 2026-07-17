@echo off
setlocal EnableDelayedExpansion
title CleanDRR Installer

:: ============================================================
::  CleanDRR Installer
::  Place this .bat file in the SAME FOLDER as:
::    CleanDRR.exe
::    users.db
::    remark_masterlist.json
::    agent_code_masterlist.json
::  Then just double-click this file.
:: ============================================================

echo.
echo  ============================================
echo    Installing CleanDRR
echo  ============================================
echo.

set "SRC=%~dp0"
set "INSTALL_DIR=%ProgramFiles%\CleanDRR"
set "DATA_DIR=%ProgramData%\CleanDRR"

:: --- 1. Make sure the installer exe is actually here ---------------------
if not exist "%SRC%CleanDRR.exe" (
    echo  ERROR: CleanDRR.exe was not found in this folder.
    echo  Make sure install_CleanDRR.bat sits next to CleanDRR.exe
    echo  before running it.
    pause
    exit /b 1
)

:: --- 2. Create install folder and copy the program ------------------------
echo  Copying program files to "%INSTALL_DIR%" ...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /Y "%SRC%CleanDRR.exe" "%INSTALL_DIR%\CleanDRR.exe" >nul

if errorlevel 1 (
    echo.
    echo  Could not write to "%INSTALL_DIR%".
    echo  Try right-clicking this .bat file and choosing "Run as administrator".
    pause
    exit /b 1
)

:: --- 3. Create shared data folder and seed it (only if not already there) -
echo  Preparing shared data folder "%DATA_DIR%" ...
if not exist "%DATA_DIR%" mkdir "%DATA_DIR%"

if not exist "%DATA_DIR%\users.db" (
    if exist "%SRC%users.db" (
        copy /Y "%SRC%users.db" "%DATA_DIR%\users.db" >nul
        echo    - users.db installed.
    )
) else (
    echo    - users.db already exists, keeping existing data.
)

if not exist "%DATA_DIR%\remark_masterlist.json" (
    if exist "%SRC%remark_masterlist.json" (
        copy /Y "%SRC%remark_masterlist.json" "%DATA_DIR%\remark_masterlist.json" >nul
        echo    - remark_masterlist.json installed.
    )
) else (
    echo    - remark_masterlist.json already exists, keeping existing data.
)

if not exist "%DATA_DIR%\agent_code_masterlist.json" (
    if exist "%SRC%agent_code_masterlist.json" (
        copy /Y "%SRC%agent_code_masterlist.json" "%DATA_DIR%\agent_code_masterlist.json" >nul
        echo    - agent_code_masterlist.json installed.
    )
) else (
    echo    - agent_code_masterlist.json already exists, keeping existing data.
)

:: --- 4. Create a Desktop shortcut ------------------------------------------
echo  Creating Desktop shortcut ...
set "VBS=%TEMP%\_shortcut_CleanDRR.vbs"

> "%VBS%" echo Set oWS = WScript.CreateObject("WScript.Shell")
>> "%VBS%" echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\CleanDRR.lnk"
>> "%VBS%" echo Set oLink = oWS.CreateShortcut(sLinkFile)
>> "%VBS%" echo oLink.TargetPath = "%INSTALL_DIR%\CleanDRR.exe"
>> "%VBS%" echo oLink.WorkingDirectory = "%INSTALL_DIR%"
>> "%VBS%" echo oLink.Description = "CleanDRR - S.P. Madrid Philippines"
>> "%VBS%" echo oLink.Save

cscript //nologo "%VBS%"
del "%VBS%"

echo.
echo  ============================================
echo    Installation complete!
echo    A "CleanDRR" shortcut was added to your Desktop.
echo  ============================================
echo.

set /p LAUNCH="Launch CleanDRR now? (Y/N): "
if /i "%LAUNCH%"=="Y" start "" "%INSTALL_DIR%\CleanDRR.exe"

pause
