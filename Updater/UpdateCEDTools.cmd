@echo off
:: Runs the PowerShell update script in a new console window
PowerShell -Command "Set-ExecutionPolicy Unrestricted -Scope Process" >> "%TEMP%\StartupLog.txt" 2>&1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.ps1"
@REM
@REM @echo off
@REM :: Launch PowerShell in a new console window and keep it open
@REM start powershell.exe -NoExit -ExecutionPolicy Bypass -File "C:\Users\Aevelina\CED_Extensions\Updater\UpdateCEDTools.ps1"
