@echo off
setlocal Enabledelayedexpansion

:: Initial setup: Maximize and Admin Check
if not "%1" == "max" start /MAX cmd /c %0 max & exit/b
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' ( goto UACPrompt ) else ( goto gotAdmin )

:UACPrompt
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs" & exit /B

:gotAdmin
if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs"
pushd "%CD%" & cd /D "%~dp0"
set "baseDir=%~dp0"

:Header
cls
color 0B
echo ============================================================
echo                          GET-OFFICE
echo ============================================================
echo.

:: Part 1: Architecture Selection
echo [ STEP 1: ARCHITECTURE ]
echo Press [Y] for 32-bit  ^|  Press [N] for 64-bit (Recommended)
choice /c yn /n /m "> Selection: "
if %errorlevel%==2 (
    set "arch=64"
    echo Selected: 64-bit
) else (
    set "arch=32"
    echo Selected: 32-bit
)

powershell Write-Host "The operation completed successfully" -ForegroundColor Green
:: Start XML Creation
echo ^<Configuration ID="genericConfig01"^> > output.xml
echo  ^<Add OfficeClientEdition="%arch%" Channel="PerpetualVL2024"^> >> output.xml
echo   ^<Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"^> >> output.xml
echo    ^<Language ID="en-gb" /^> >> output.xml

echo.
echo ============================================================
echo [ STEP 2: APP SELECTION ]
echo Press [Y] to INSTALL  ^|  Press [N] to EXCLUDE
echo ============================================================
echo.

:: App selection loop
call :AskApp "Access" "Access"
call :AskApp "Excel" "Excel"
call :AskApp "Skype/Lync" "Lync"
call :AskApp "OneDrive" "OneDrive"
call :AskApp "OneNote" "OneNote"
call :AskApp "Outlook" "Outlook"
call :AskApp "PowerPoint" "PowerPoint"
call :AskApp "Publisher" "Publisher"
call :AskApp "Word" "Word"

:: Close XML structure
echo   ^</Product^> >> output.xml
echo  ^</Add^> >> output.xml
echo  ^<Property Name="SharedComputerLicensing" Value="0" /^> >> output.xml
echo  ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /^> >> output.xml
echo  ^<Property Name="DeviceBasedLicensing" Value="0" /^> >> output.xml
echo  ^<Property Name="SCLCacheOverride" Value="0" /^> >> output.xml
echo  ^<Property Name="AUTOACTIVATE" Value="1" /^> >> output.xml
echo  ^<Updates Enabled="TRUE" /^> >> output.xml
echo  ^<AppSettings^> >> output.xml
echo   ^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^> >> output.xml
echo   ^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^> >> output.xml
echo   ^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^> >> output.xml
echo  ^</AppSettings^> >> output.xml
echo ^</Configuration^> >> output.xml

echo.
echo ============================================================
echo [ STEP 3: DEPLOYMENT ]
echo ============================================================
if not exist "temp" mkdir "temp"
powershell -Command "New-Item -ItemType Directory -Force '%~dp0temp' | Out-Null; Invoke-WebRequest -Uri 'https://reubak1.github.io/Get-Office/officedeploymenttool_18129-20030.exe' -OutFile '%~dp0temp\officedeploymenttool_18129-20030.exe'"
move "output.xml" "temp" >nul
cd temp


echo Extracting deployment files...
officedeploymenttool_18129-20030.exe /quiet /extract:"%CD%"

echo.
echo ============================================================
echo [ Ready to install ]
echo ============================================================
pause
echo [!] Installing Office. This window will remain open...
setup.exe /configure output.xml

echo.
echo All credit to massgrave.dev for the activation method.
echo [!] Running Activation...
powershell -ExecutionPolicy Bypass -NoProfile -Command " & ([ScriptBlock]::Create((irm https://reubak1.github.io/Get-Office/activation))) /Ohook"


:: CLEANUP SECTION
echo.
echo [!] Cleaning up temporary files...
cd /D "%baseDir%"
timeout /t 2 >nul
rd /s /q "%baseDir%temp"

echo.
echo ============================================================
echo        INSTALLATION COMPLETE. YOU CAN CLOSE THIS WINDOW.
echo ============================================================
pause
exit

:: Function to handle the auto-accept choice
:AskApp
choice /c yn /n /m "Install %~1? (y/n): "
if %errorlevel%==2 (
powershell Write-Host "The operation completed successfully" -ForegroundColor Green
    echo    ^<ExcludeApp ID="%~2" /^> >> output.xml
) else (
powershell Write-Host "The operation completed successfully" -ForegroundColor Green
)
exit /b
