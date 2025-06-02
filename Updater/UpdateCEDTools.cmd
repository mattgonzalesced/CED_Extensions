@REM @echo off
@REM :: Runs the PowerShell update script in a new console window
@REM PowerShell -Command "Set-ExecutionPolicy Unrestricted -Scope Process" >> "%TEMP%\StartupLog.txt" 2>&1
@REM powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.ps1"

@echo off
:: Launch PowerShell in a new console window and keep it open
start powershell.exe -NoExit -ExecutionPolicy Bypass -File "C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.ps1"
