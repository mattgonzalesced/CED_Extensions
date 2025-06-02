Write-Output "=== Updating or Cloning CED Tools Extensions in pyRevit Extensions folder ==="

$extensionsDir = "$env:APPDATA\pyRevit\Extensions"
$repoDir = Join-Path -Path $extensionsDir -ChildPath "CED_Extensions"
$gitUrl = "https://github.com/mattgonzalesced/CED_Extensions.git"
$branchName = "develop"

Write-Output "Repo directory: $repoDir"
Write-Output "Repo URL: $gitUrl"
Write-Output "Branch: $branchName"

if (-not (Test-Path $repoDir)) {
    Write-Output "Repo folder does not exist. Cloning..."

    # IMPORTANT: Go to the PARENT folder where clone will create the CED_Extensions folder
    Set-Location $extensionsDir
    $cloneResult = git clone -b $branchName $gitUrl 2>&1
    Write-Output $cloneResult

    if (-not (Test-Path $repoDir)) {
        Write-Output "❌ Clone failed! Repo dir still missing."
        exit 1
    } else {
        Write-Output "✅ Clone complete."
    }
} else {
    Write-Output "Repo folder exists. Pulling updates..."

    # Now that it exists, we can go into it
    Set-Location $repoDir
    $pullResult = git pull origin $branchName 2>&1
    Write-Output $pullResult
    Write-Output "✅ Updates pulled."
}

Write-Output "=== Update/Clone complete! Please reload pyRevit in Revit. ==="
