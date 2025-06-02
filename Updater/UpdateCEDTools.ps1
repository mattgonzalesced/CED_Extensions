# UpdateCEDTools.ps1
# Actually pulls updates and sets up extension paths

Write-Host "=== Updating CED Tools Extensions ==="

$repoDir = "$env:APPDATA\pyRevit\Extensions\CED_Extensions"
$branchName = "main"

Set-Location $repoDir
git pull origin $branchName

# Re-add extension path (safe to re-run)
pyrevit extensions paths add $repoDir

# Enable rocketmode and telemetry
pyrevit configs rocketmode enable
pyrevit configs telemetry disable

# Remove the following line after verification
pyrevit extensions paths add "C:\Users\Aevelina\CED_Extensions"

Write-Host "=== Update complete! Please reload pyRevit in Revit. ==="
