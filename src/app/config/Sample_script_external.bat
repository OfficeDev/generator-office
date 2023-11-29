@echo off
setlocal enabledelayedexpansion

where /q git
if ErrorLevel 1 (
    echo Git is not installed, installing now...
    powershell -Command "Invoke-WebRequest https://github.com/git-for-windows/git/releases/download/v2.33.0.windows.2/Git-2.33.0.2-64-bit.exe -OutFile git-installer.exe"
    start /wait git-installer.exe /VERYSILENT
    del git-installer.exe
    echo Git has been installed.
    echo Restarting script after installed git...
    @REM call %~dpnx0
    @REM exit
) else (
    echo Git is already installed!
)


where /q node
if ErrorLevel 1 (
    echo Node.js is not installed, installing now...
    powershell -Command "Invoke-WebRequest https://nodejs.org/dist/v20.9.0/node-v20.9.0-x64.msi -OutFile node.msi"
    msiexec /i node.msi /passive
    del node.msi
    echo Node.js has been installed.
    @REM echo Restarting script after installed node.js...
    @REM call %~dpnx0
    @REM exit
) else (
    echo Node.js is already installed!
)

@REM call %~dpnx0
@REM exit


@REM Now Node.js, git have all prepared. Install Yeoman Office.

echo Git and Node.js prepared. Checking Yeoman Office...
yo --generators | findstr /C:"office"
if ErrorLevel 1 (
    echo Yeoman Office is not installed, installing now...
    npm install -g yo generator-office
    echo Yeoman Office has been installed.
) else (
    echo Yeoman Office has already been installed.
)

@REM Now Yeoman Office has been installed. Create a sample project.

set foldername=Office_sample
set /a counter=0

:loop
if exist %foldername% (
    set /a counter=counter + 1
    set foldername=Office_sample_%counter%
    goto loop
)

yo office --output %foldername% --projectType excel_sample

@REM echo Sample script has been finished.
@REM exit

pause