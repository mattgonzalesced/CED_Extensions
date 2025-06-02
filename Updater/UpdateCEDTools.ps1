[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Output "=== Starting CED Tools Update/Clone ==="

$repoUrl = "https://github.com/mattgonzalesced/CED_Extensions.git"
$extensionsDir = "$env:APPDATA\pyRevit\Extensions"
$repoDir = Join-Path -Path $extensionsDir -ChildPath "CED_Extensions"
$branchName = "develop"

if (-not (Test-Path $repoDir)) {
    Write-Output "Repo folder does not exist. Cloning repo to Extensions folder..."
    Set-Location $extensionsDir
    git clone -b $branchName $repoUrl CED_Extensions 2>&1
    if (-not (Test-Path $repoDir)) {
        Write-Output "❌ Clone failed! Repo dir still missing."
        exit 1
    }
    Write-Output "✅ Clone complete!"
} else {
    Write-Output "Repo folder exists. Pulling updates..."
    Set-Location $repoDir
    git pull origin $branchName 2>&1
    Write-Output "✅ Updates pulled."
}

Write-Output "Adding repo folder as an extension search path..."
pyrevit extensions paths add $repoDir 2>&1
Write-Output "✅ Extension search path added."

Write-Output "=== Update/Clone complete! Please manually reload pyRevit in Revit. ==="

