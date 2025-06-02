@echo off
:: Runs the PowerShell update script in a new console window
PowerShell -Command "Set-ExecutionPolicy Unrestricted -Scope Process" >> "%TEMP%\StartupLog.txt" 2>&1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Aevelina\AppData\Roaming\pyRevit\Extensions\CED_Extensions\Updater\UpdateCEDTools.ps1"

