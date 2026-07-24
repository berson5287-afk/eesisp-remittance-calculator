@echo off
setlocal enabledelayedexpansion
title Build EESISP Remittance Calculator
cd /d "%~dp0"

set "SRC=eesisp_calculator.py"
set "APPNAME=EESISP_Calculator"

echo(
echo  ==========================================================
echo    Building %APPNAME%.exe
echo  ==========================================================
echo(

rem ---------------------------------------------------------------
rem  1. the source file has to sit next to this script
rem ---------------------------------------------------------------
if not exist "%SRC%" (
    echo  ERROR: %SRC% was not found in this folder.
    echo(
    echo  Put build_exe.bat in the SAME folder as %SRC%
    echo  and run it again.
    goto :fail
)

rem ---------------------------------------------------------------
rem  2. find Python
rem ---------------------------------------------------------------
set "PYCMD="
py -3 --version >nul 2>&1
if not errorlevel 1 set "PYCMD=py -3"

if not defined PYCMD (
    python --version >nul 2>&1
    if not errorlevel 1 set "PYCMD=python"
)

if not defined PYCMD (
    echo  ERROR: Python was not found on this PC.
    echo(
    echo  Install it from  https://www.python.org/downloads/
    echo  and tick "Add python.exe to PATH" on the first setup screen.
    goto :fail
)

for /f "tokens=*" %%v in ('%PYCMD% --version 2^>^&1') do set "PYVER=%%v"
echo  Python:      !PYVER!

rem ---------------------------------------------------------------
rem  3. tkinter has to be present (it ships with the python.org build)
rem ---------------------------------------------------------------
%PYCMD% -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo(
    echo  ERROR: this Python has no tkinter.
    echo  Re-run the Python installer, choose Modify, and make sure
    echo  "tcl/tk and IDLE" is ticked.
    goto :fail
)
echo  tkinter:     found

rem ---------------------------------------------------------------
rem  4. PyInstaller
rem ---------------------------------------------------------------
echo(
echo  Installing / updating PyInstaller...
%PYCMD% -m pip install --upgrade --quiet pip
%PYCMD% -m pip install --upgrade --quiet pyinstaller
if errorlevel 1 (
    echo(
    echo  ERROR: PyInstaller could not be installed.
    echo  If this PC is behind a proxy, try:
    echo     %PYCMD% -m pip install pyinstaller
    echo  and read the message it prints.
    goto :fail
)

rem ---------------------------------------------------------------
rem  5. clean out anything from a previous run
rem ---------------------------------------------------------------
echo  Cleaning previous build...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "%APPNAME%.spec" del /q "%APPNAME%.spec"

rem ---------------------------------------------------------------
rem  6. build
rem ---------------------------------------------------------------
echo(
echo  Building. This takes a minute or two - leave the window open.
echo(
%PYCMD% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --clean ^
    --noconfirm ^
    --name "%APPNAME%" ^
    "%SRC%"

if errorlevel 1 goto :buildfail
if not exist "dist\%APPNAME%.exe" goto :buildfail

rem ---------------------------------------------------------------
rem  7. done
rem ---------------------------------------------------------------
echo(
echo  ==========================================================
echo    Finished.
echo(
echo    Your program:
echo      %CD%\dist\%APPNAME%.exe
echo(
echo    Copy that single file anywhere you like - onto the work PC,
echo    a network share, a USB stick. Nothing else needs installing.
echo(
echo    Your employee list, rates and saved periods live in
echo      %%APPDATA%%\EESISP Calculator\eesisp_data.json
echo    NOT inside the .exe, so rebuilding never wipes your data.
echo(
echo    First launch is slow for a second or two while the one-file
echo    exe unpacks itself. Windows SmartScreen may also warn about
echo    an unknown publisher, because the file is not code-signed:
echo    choose "More info" then "Run anyway".
echo  ==========================================================
echo(

choice /c YN /n /m "  Open the dist folder now? [Y/N] "
if errorlevel 2 goto :done
start "" "%CD%\dist"
goto :done

rem ---------------------------------------------------------------
:buildfail
echo(
echo  ==========================================================
echo    BUILD FAILED
echo  ==========================================================
echo(
echo  Scroll up - PyInstaller prints the reason near the end.
echo  The usual causes are antivirus locking the dist folder, or
echo  running this from a synced folder (OneDrive/Dropbox).
echo  Try copying both files to C:\Temp\eesisp and running again.
echo(
goto :fail

:fail
echo(
pause
exit /b 1

:done
echo(
pause
exit /b 0
