@echo off
:: Update CED_Extensions and refresh environment

:: Path to the local clone
set "REPO_DIR=C:\Program Files\CEDApp\CED_Extensions"

:: 1) Pull latest updates
cd /d "%REPO_DIR%"
git pull origin main

:: 2) Add this path as an extension path (safe to re-run)
pyrevit extensions paths add "%REPO_DIR%"

:: (Optional) Enable rocketmode and telemetry
pyrevit configs rocketmode enable
pyrevit configs telemetry enable

echo.
echo Update complete! Please manually reload pyRevit in Revit to activate updates.
pause
