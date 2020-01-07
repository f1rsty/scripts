@echo off


del /F /S /Q  %temp%\*.* 
del /F /S /Q  C:\Windows\Prefetch\*.* 
del /F /S /Q  C:\Windows\Temp\*.* 
del /F /S /Q  %localappdata%\Microsoft\Windows\History\*.* 
del /F /S /Q  %localappdata%\Microsoft\Windows\WebCache\*.* 
del /F /S /Q  "%localappdata%\Microsoft\Windows\Temporary Internet Files"\*.* 
del /F /S /Q  %appdata%\Microsoft\Windows\Cookies\Low\*.* 
del /F /S /Q  %appdata%\Microsoft\Windows\Cookies\*.* 
