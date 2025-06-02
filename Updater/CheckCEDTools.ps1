$repoDir = "$env:APPDATA\pyRevit\Extensions\CED_Extensions"

if (-not (Test-Path $repoDir)) {
    Write-Output "status: clone-needed"
    exit
}

# Use pyRevit CLI to check for updates
$updateResult = pyrevit extensions update $repoDir 2>&1
if ($updateResult -match "Already up to date.") {
    Write-Output "status: up-to-date"
} else {
    Write-Output "status: updates-available"
}
