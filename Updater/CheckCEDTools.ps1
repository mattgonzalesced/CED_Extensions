$repoDir = "$env:APPDATA\pyRevit\Extensions\CED_Extensions"
$branchName = "develop"

if (-not (Test-Path $repoDir)) {
    Write-Output "status: clone-needed"
    exit
}

Set-Location $repoDir
git fetch origin $branchName

$localHash = git rev-parse HEAD
$remoteHash = git rev-parse origin/$branchName

if ($localHash -eq $remoteHash) {
    Write-Output "status: up-to-date"
} else {
    Write-Output "status: updates-available"
}
