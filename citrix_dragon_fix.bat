@echo off

:: Fixes Dragon "Citrix component channels missing"
:: Error occurs due to newest Citrix receiver without doing a vsync restore

if "%~s0"=="%~s1" ( cd %~sp1 & shift ) else (
  echo CreateObject^("Shell.Application"^).ShellExecute "%~s0","%~0 %*","","runas",1 >"%tmp%%~n0.vbs" & "%tmp%%~n0.vbs" & del /q "%tmp%%~n0.vbs" & goto :eof
)
:eof

taskkill /F /IM "Receiver.exe" /T
taskkill /F /IM "SelfServicePlugin.exe" /T
taskkill /F /IM "wfcrun32.exe" /T
taskkill /F /IM "SelfService.exe" /T
taskkill /F /IM "concentr.exe" /T
taskkill /F /IM "natspeak.exe" /T

dserv="\\MHC-MSASNDVNMS1\Dragon Software\Dragon Medical 360 Client Software\12.51.217.164_DMNE_2.7.5_FullClient_ENU\DNS12_DVD1\vSyncRestorationPatch"
net use Q: %dserv%
pushd Q:
start .\vSyncRestorer.exe
net use Q: /D /y
