# UpdateCEDTools.ps1
# PowerShell script to deploy and register CED Extensions for pyRevit (develop branch)

Write-Host "=== Starting CED Tools Update ==="

# Set paths
$pyRevitExtensionsDir = "$env:APPDATA\pyRevit\Extensions"
$repoDir = Join-Path -Path $pyRevitExtensionsDir -ChildPath "CED_Extensions"
$gitUrl = "https://github.com/mattgonzalesced/CED_Extensions.git"
$branchName = "develop"

# Create Extensions folder if needed
if (-not (Test-Path $pyRevitExtensionsDir)) {
    Write-Host "Creating pyRevit Extensions folder..."
    New-Item -Path $pyRevitExtensionsDir -ItemType Directory -Force
}

# Clone or update the repo for the develop branch
if (-not (Test-Path $repoDir)) {
    Write-Host "Cloning CED_Extensions (develop branch) to pyRevit Extensions..."
    git clone --branch $branchName $gitUrl $repoDir
} else {
    Write-Host "Repo exists, switching to develop branch and pulling updates..."
    Set-Location $repoDir
    git fetch origin
    git checkout $branchName
    git pull origin $branchName
}

# Show contents
Write-Host "Repo contents:"
Get-ChildItem $repoDir

# Add this repo as an extension search path (safe to re-run)
Write-Host "Adding CED_Extensions path to pyRevit search paths..."
pyrevit extensions paths add $repoDir

# Enable rocketmode and telemetry
Write-Host "Enabling rocketmode and telemetry..."
pyrevit configs rocketmode enable
pyrevit configs telemetry disable

# Show final pyRevit environment
Write-Host "Showing final pyRevit environment:"
pyrevit env

Write-Host "=== Update complete! Please reload pyRevit in Revit. ==="
Pause
