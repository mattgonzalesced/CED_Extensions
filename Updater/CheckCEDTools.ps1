# CheckCEDTools.ps1
# Only checks repo status, does NOT make changes

Write-Host "=== Checking CED Tools Extension Status ==="

$repoDir = "$env:APPDATA\pyRevit\Extensions\CED_Extensions"
$gitUrl = "https://github.com/mattgonzalesced/CED_Extensions.git"
$branchName = "develop"

# Check if repo exists
if (-not (Test-Path $repoDir)) {
    Write-Host "status: clone-needed"
    exit
}

# Check for changes
Set-Location $repoDir
git fetch origin $branchName

$localHash = git rev-parse HEAD
$remoteHash = git rev-parse origin/$branchName

if ($localHash -eq $remoteHash) {
    Write-Host "status: up-to-date"
    exit
} else {
    Write-Host "status: updates-available"
    exit
}
