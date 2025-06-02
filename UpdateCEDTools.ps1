# UpdateCEDTools.ps1
# PowerShell script to deploy and register CED Extensions for pyRevit

Write-Host "=== Starting CED Tools Update ==="

# Set paths
$repoDir = "C:\Program Files\CEDApp\CED_Extensions"
$gitUrl = "https://github.com/mattgonzalesced/CED_Extensions.git"

# Create CEDApp folder if it doesn't exist
if (-not (Test-Path "C:\Program Files\CEDApp")) {
    Write-Host "Creating CEDApp folder..."
    New-Item -Path "C:\Program Files\CEDApp" -ItemType Directory -Force
}

# Clone or update the repo
if (-not (Test-Path $repoDir)) {
    Write-Host "Cloning repo to Program Files..."
    git clone $gitUrl $repoDir
} else {
    Write-Host "Repo exists, pulling updates..."
    Set-Location $repoDir
    git pull origin main
}

# Show the contents of the repo
Write-Host "Repo contents:"
Get-ChildItem $repoDir

# Register each extension
$extensions = @("AE PyDev.extension", "AE pyTools.extension", "WM Tools.extension")
foreach ($ext in $extensions) {
    $jsonPath = Join-Path $repoDir $ext "extension.json"
    if (Test-Path $jsonPath) {
        Write-Host "Registering $ext with pyRevit..."
        pyrevit extensions sources add $jsonPath
    } else {
        Write-Host "Missing extension.json for $ext"
    }
}

# Enable rocketmode and telemetry
Write-Host "Enabling rocketmode and telemetry..."
pyrevit configs rocketmode enable
pyrevit configs telemetry enable

# Show final pyRevit environment
Write-Host "Showing final pyRevit environment:"
pyrevit env

Write-Host "=== Update complete! Please reload pyRevit in Revit ==="
Pause
