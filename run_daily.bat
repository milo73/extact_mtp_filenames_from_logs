@echo off
:: MailToPost daily log processor
:: Scheduled via Windows Task Scheduler – runs each morning after midnight.
::
:: Task Scheduler settings (recommended):
::   General  : Run whether user is logged on or not
::   Triggers : Daily, e.g. 06:00 AM
::   Actions  : Start a program
::                Program : C:\Windows\System32\cmd.exe
::                Arguments: /c "F:\Planetpress\MailToPost\Scripts\run_daily.bat"
::   Settings : If the task fails, restart every 5 minutes, up to 3 times

::
:: IMPORTANT: Set the Azure storage key as a Windows system environment variable
:: so it is never stored in this file. In an elevated PowerShell prompt run:
::   [System.Environment]::SetEnvironmentVariable('AZURE_STORAGE_KEY','<your-key>','Machine')
:: Then restart the Task Scheduler service so it picks up the new variable.
::
set PYTHON="C:\Program Files\Python313\python.exe"
set SCRIPT="F:\Planetpress\MailToPost\Scripts\run_daily.py"
set CONFIG="F:\Planetpress\MailToPost\Scripts\config.ini"
set LOGDIR="F:\Planetpress\MailToPost\Logs"

:: Create log dir if it doesn't exist
if not exist %LOGDIR% mkdir %LOGDIR%

echo [%date% %time%] Starting MailToPost daily run >> %LOGDIR%\run_daily.log

:: ---------------------------------------------------------------------------
:: Ensure the Azure file share is reachable and mapped
:: ---------------------------------------------------------------------------
powershell.exe -NonInteractive -NoProfile -Command ^
  "$result = Test-NetConnection -ComputerName stddifilepweu02.file.core.windows.net -Port 445; ^
   if ($result.TcpTestSucceeded) { ^
     cmd.exe /C ('cmdkey /add:\"stddifilepweu02.file.core.windows.net\" /user:\"localhost\stddifilepweu02\" /pass:\"' + $env:AZURE_STORAGE_KEY + '\"'); ^
     if (-not (Test-Path 'Z:\')) { ^
       New-PSDrive -Name Z -PSProvider FileSystem -Root '\\stddifilepweu02.file.core.windows.net\ddi-data-p-01' -Persist | Out-Null; ^
       Write-Host 'Drive Z: mapped.' ^
     } else { ^
       Write-Host 'Drive Z: already mapped.' ^
     } ^
   } else { ^
     Write-Error 'Cannot reach stddifilepweu02.file.core.windows.net on port 445. Check firewall or VPN.'; ^
     exit 1 ^
   }" >> %LOGDIR%\run_daily.log 2>&1

if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] Network mapping failed – aborting. >> %LOGDIR%\run_daily.log
    exit /b %ERRORLEVEL%
)

:: ---------------------------------------------------------------------------
:: Run the daily parser and e-mailer
:: ---------------------------------------------------------------------------
%PYTHON% %SCRIPT% --config %CONFIG% >> %LOGDIR%\run_daily.log 2>&1

if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] run_daily.py failed with exit code %ERRORLEVEL% >> %LOGDIR%\run_daily.log
    exit /b %ERRORLEVEL%
)

echo [%date% %time%] Done. >> %LOGDIR%\run_daily.log
