:: INSTALL SCRIPT BELOW
@echo off

:: Stop Known Conflicting Applications
taskkill /IM Outlook.exe /F
taskkill /IM WinWord.exe /F
taskkill /IM natspeak.exe /F
taskkill /IM excel.exe /F
taskkill /IM pnamain.exe /F
taskkill /IM wfcrun32.exe /F
taskkill /IM wfica32.exe /F
taskkill /IM dgnuiasvr.exe /F

:: Remove all traces of the cache files from the local computer
set "folder="
for /D %%b in ("C:\Users\*") do (
    for /D %%c in ("%%~b\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\*") do (
            rmdir /S /Q %%c
        )
    )
)

rmdir /S /Q "C:\ProgramData\Nuance\NaturallySpeaking12\Users\*"
rmdir /S /Q "C:\ProgramData\Nuance\NaturallySpeaking12\RoamingUsers\*"
rmdir /S /Q "C:\ProgramData\Nuance\NaturallySpeaking12\BadUsers\*"

:: Setup Install Directory where you have your installs
set appdir="\\mhc-msasndvnms1.mclaren.org\Dragon Software\Dragon Medical 360 Client Software\12.51.217.164_DMNE_2.7.5_FullClient_ENU\DNS12_DVD1"

:: Mount NAS to the Q: Drive
net use q: %appdir%

:: Push the directory to the Q: Drive (Which points to the NAS Drive)
pushd q:

: InstallDragon
setup.exe /s /v"NAS_ADDRESS=MHC-MSASNDVNMS1 SERIALNUMBER=A709A-K13-E33D-NR9A-D1 /qn"
::/l*xv C:\Files\InstallDragon12.log"
IF NOT "%ERRORLEVEL%"=="0" EXIT /B %ERRORLEVEL%

:: Remove old custom Dragon 9.5 link
::del /Q /F "%ALLUSERSPROFILE%\Desktop\Dragon NaturallySpeaking 9.5.lnk"

:: Replace Custom Application Shortcuts
::xcopy Network Edition.lnk "%ALLUSERSPROFILE%\Desktop" /i /y
::xcopy "%appdir%\Network Edition.lnk" "%ALLUSERSPROFILE%\Start Menu\Programs\Dragon Medical 12" /i /y

::xcopy "%appdir%\dragon.bat "%ALLUSERSPROFILE%\Start Menu\Programs\Startup" /i /y
