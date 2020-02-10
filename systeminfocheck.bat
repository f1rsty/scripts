:: Set location for logs
set LOGPATH=%SystemDrive%\Logs
set LOGFILE=%COMPUTERNAME%_system_information.log

:::::::::::::::::::::
:: PREP AND CHECKS ::
:::::::::::::::::::::
@echo off && cls
set SCRIPT_VERSION=0.1
set SCRIPT_UPDATED=2-10-2020

:: Get the date into ISO 8601 standard format (yyyy-mm-dd) so we can use it
FOR /f %%a in ('WMIC OS GET LocalDateTime ^| find "."') DO set DTS=%%a
set CUR_DATE=%DTS:~0,4%-%DTS:~4,2%-%DTS:~6,2%

title System Checker v%SCRIPT_VERSION% (%SCRIPT_UPDATED%)

if not exist %LOGPATH% mkdir %LOGPATH%
if exist "%LOGPATH%\%LOGFILE%" del "%LOGPATH%\%LOGFILE%"

call :log "%CUR_DATE% %TIME%   Targeting individual JRE versions..."

:::::::::::::::
:: FUNCTIONS ::
:::::::::::::::
:log
echo:%~1 >> "%LOGPATH%\%LOGFILE%"
echo:%~1
EXIT /B 0
