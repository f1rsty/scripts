@ECHO OFF
if "%~s0"=="%~s1" ( cd %~sp1 & shift ) else (
  echo CreateObject^("Shell.Application"^).ShellExecute "%~s0","%~0 %*","","runas",1 >"%tmp%%~n0.vbs" & "%tmp%%~n0.vbs" & del /q "%tmp%%~n0.vbs" & goto :eof
)
:eof

:MENU
CLS
ECHO.
ECHO =================================================
ECHO Please select where the computer will be located.
ECHO =================================================
ECHO.
ECHO 1 - Nurses Station
ECHO 2 - Bed Side
ECHO 3 - Exit
ECHO.
SET /P M=Enter your selection: 
IF %M%==1 SET STATION=OBIX-NS
IF %M%==2 SET STATION=OBIX-BS
IF %M%==3 GOTO EOF

CALL :MODIFYREGISTRY
CALL :DOTNETCHECK
CALL :SETPATH
CALL :ADDSHORTCUT
CALL :IMPORTCERT
EXIT /B 0

:DOTNETCHECK
@echo off
dism /online /get-features /format:table | find /i "NetFx3" | find /v "Microsoft" | find "Enabled" >nul
if %ERRORLEVEL% == 1 (
	dism /online /Enable-Feature /FeatureName:NetFx3 /NoRestart
    echo.
) else (
	echo NetFx3 is already enabled.
	PING localhost -n 3 >NUL
)
echo.
EXIT /B 0

:ADDSHORTCUT
@echo off
set SCRIPT1="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"
set SCRIPT2="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%.vbs"

:: Creates the Patient Manager shortcut
echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT1%
echo sLinkFile = "C:\Users\Public\Desktop\Patient Manager.lnk" >> %SCRIPT1%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT1%
echo oLink.WindowStyle = 0 >> %SCRIPT1%
echo oLink.TargetPath = "%PROGRAMFILES%\CCSI\ptmgr.exe" >> %SCRIPT1%
echo oLink.Save >> %SCRIPT1%

:: Creates the Surveillance shortcut
echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT2%
echo sLinkFile = "C:\Users\Public\Desktop\Surveillance.lnk" >> %SCRIPT2%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT2%
echo oLink.WindowStyle = 0 >> %SCRIPT2%
echo oLink.TargetPath = "%PROGRAMFILES%\CCSI\svdsp.exe" >> %SCRIPT2%
echo oLink.Save >> %SCRIPT2%

:: Calls the scripts created from above
cscript /nologo %SCRIPT1%
cscript /nologo %SCRIPT2%

:: Cleanup task for above scripts
del %SCRIPT1%
del %SCRIPT2%


:: Move the shortcuts to the start up location
move "C:\Users\Public\Desktop\Surveillance.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
move "C:\Users\Public\Desktop\Patient Manager.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"

EXIT /B 0
