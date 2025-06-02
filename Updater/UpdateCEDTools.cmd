Write-Output "=== Starting CED Tools Update using pyRevit CLI extend ==="

$extensionName = "CED_Extensions"
$repoUrl = "https://github.com/mattgonzalesced/CED_Extensions.git"
$extensionsDir = "$env:APPDATA\pyRevit\Extensions"
$repoDir = Join-Path -Path $extensionsDir -ChildPath $extensionName
$branchName = "develop"

if (-not (Test-Path $repoDir)) {
    Write-Output "Extension folder does not exist. Cloning with pyRevit CLI..."
    pyrevit extend ui $extensionName $repoUrl --dest="$repoDir" --branch=$branchName 2>&1
    if (-not (Test-Path $repoDir)) {
        Write-Output "❌ Clone failed! Repo dir still missing."
        exit 1
    }
    Write-Output "✅ Clone complete!"
} else {
    Write-Output "Extension folder exists. Updating with pyRevit CLI..."
    pyrevit extensions update $extensionName 2>&1
    Write-Output "✅ Updates pulled."
}

Write-Output "Re-adding extension path to pyRevit (safe to re-run)..."
pyrevit extensions paths add $repoDir 2>&1

Write-Output "=== Update/Clone complete! Please manually reload pyRevit in Revit. ==="
