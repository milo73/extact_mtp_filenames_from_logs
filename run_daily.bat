@echo off
:: MailToPost daily log processor
:: Scheduled via Windows Task Scheduler – runs each morning after midnight.
::
:: Task Scheduler settings (recommended):
::   General  : Run whether user is logged on or not
::   Triggers : Daily, e.g. 06:00 AM
::   Actions  : Start a program
::                Program : "C:\Program Files\Python313\python.exe"
::                Arguments: F:\Planetpress\MailToPost\Scripts\run_daily.bat
::   Settings : If the task fails, restart every 5 minutes, up to 3 times

set PYTHON="C:\Program Files\Python313\python.exe"
set SCRIPT="F:\Planetpress\MailToPost\Scripts\run_daily.py"
set CONFIG="F:\Planetpress\MailToPost\Scripts\config.ini"
set LOGDIR="F:\Planetpress\MailToPost\Logs"

:: Create log dir if it doesn't exist
if not exist %LOGDIR% mkdir %LOGDIR%

:: Run and append stdout + stderr to a local log file
%PYTHON% %SCRIPT% --config %CONFIG% >> %LOGDIR%\run_daily.log 2>&1

if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] run_daily.py failed with exit code %ERRORLEVEL% >> %LOGDIR%\run_daily.log
    exit /b %ERRORLEVEL%
)
