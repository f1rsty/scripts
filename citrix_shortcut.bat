@echo off

:: Set the variables
set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"
set SHORTCUTLINK="https://mhccitrixapps.mclaren.org"
set LINKTEXT=McLaren WSP Citrix

:: Create a vbs script from this batch file
echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%
echo sLinkFile = "%USERPROFILE%\Desktop\%LINKTEXT%.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.Arguments = %SHORTCUTLINK% >> %SCRIPT%
echo oLink.WindowStyle = 0 >> %SCRIPT%
echo oLink.TargetPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" >> %SCRIPT%
echo oLink.Save >> %SCRIPT%

cscript /nologo %SCRIPT%

del %SCRIPT%

EXIT /b 0
