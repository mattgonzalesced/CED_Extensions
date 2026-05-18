[CmdletBinding()]
param(
    [string]$Version = "0.0.1.0"
)

$ErrorActionPreference = "Stop"
$ScriptDir  = $PSScriptRoot
$RepoRoot   = Split-Path $ScriptDir -Parent
$StagingDir = Join-Path $ScriptDir "staging"
$OutputDir  = Join-Path $ScriptDir "output"

$ManagedFolders = @(
    "AE pyTools.extension",
    "CED ElecTools.extension",
    "CED MechTools.extension",
    "CEDLib.lib",
    "CED_pyTelemetry",
    "WM Tools.extension"
)

Write-Host "=== Building Updater_pyrevit.msi v$Version ===" -ForegroundColor Cyan

# Clean staging + output
if (Test-Path $StagingDir) { Remove-Item $StagingDir -Recurse -Force }
if (Test-Path $OutputDir)  { Remove-Item $OutputDir  -Recurse -Force }
New-Item $StagingDir -ItemType Directory | Out-Null
New-Item $OutputDir  -ItemType Directory | Out-Null

# Stage managed folders from repo root
Write-Host "`n[1/4] Staging extension folders..." -ForegroundColor Yellow
foreach ($folder in $ManagedFolders) {
    $src = Join-Path $RepoRoot $folder
    if (Test-Path $src) {
        Copy-Item $src $StagingDir -Recurse
        Write-Host "  + $folder"
    } else {
        Write-Warning "  ! $folder NOT FOUND in repo root, skipping"
    }
}

# Harvest with heat.exe
# -ag: auto-generate component GUIDs at compile time (stable across builds, derived from install path)
Write-Host "`n[2/4] Harvesting file list..." -ForegroundColor Yellow
$ExtensionsWxs = Join-Path $ScriptDir "Extensions.wxs"
& heat.exe dir $StagingDir `
    -ag -sfrag -srd -scom -sreg `
    -cg ExtensionsGroup `
    -dr APPDATAPYREVITEXT `
    -var var.StagingDir `
    -out $ExtensionsWxs
if ($LASTEXITCODE -ne 0) { throw "heat.exe failed (exit $LASTEXITCODE)" }

# Compile + link
Push-Location $ScriptDir
try {
    Write-Host "`n[3/4] Compiling .wxs -> .wixobj..." -ForegroundColor Yellow
    & candle.exe -nologo -ext WixUtilExtension -ext WixUIExtension `
        -dStagingDir="$StagingDir" `
        -dProductVersion="$Version" `
        "Product.wxs" "Extensions.wxs"
    if ($LASTEXITCODE -ne 0) { throw "candle.exe failed (exit $LASTEXITCODE)" }

    Write-Host "`n[4/4] Linking .wixobj -> .msi..." -ForegroundColor Yellow
    $OutputMsi = Join-Path $OutputDir "Updater_pyrevit.msi"
    & light.exe -nologo -ext WixUtilExtension -ext WixUIExtension `
        -sw1076 -sice:ICE38 -sice:ICE43 -sice:ICE57 -sice:ICE61 -sice:ICE64 -sice:ICE91 `
        -cultures:en-us `
        -out $OutputMsi `
        "Product.wixobj" "Extensions.wixobj"
    if ($LASTEXITCODE -ne 0) { throw "light.exe failed (exit $LASTEXITCODE)" }
}
finally {
    Pop-Location
}

$OutputMsi = Join-Path $OutputDir "Updater_pyrevit.msi"
$msiInfo = Get-Item $OutputMsi
Write-Host "`nBuilt: $OutputMsi" -ForegroundColor Green
Write-Host "Size:  $([math]::Round($msiInfo.Length / 1MB, 2)) MB"
