$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$metaJson = python -c "from metadata import APP_NAME, APP_VERSION; import json; print(json.dumps({'name': APP_NAME, 'version': APP_VERSION}, ensure_ascii=False))"
if (-not $metaJson) {
    throw "Failed to read metadata.py"
}

$meta = $metaJson | ConvertFrom-Json
$appName = [string]$meta.name
$appVersion = [string]$meta.version
$packageName = "$appName v$appVersion"

$distDir = Join-Path $PSScriptRoot "dist"
$buildDir = Join-Path $PSScriptRoot "build"
$pyinstallerDir = Join-Path $distDir $appName
$releaseRoot = Join-Path $PSScriptRoot "release"
$releaseDir = Join-Path $releaseRoot $packageName
$zipPath = Join-Path $releaseRoot ($packageName + ".zip")

Write-Host "Building $packageName ..." -ForegroundColor Cyan

if (Test-Path $buildDir) { Remove-Item -Recurse -Force $buildDir }
if (Test-Path $pyinstallerDir) { Remove-Item -Recurse -Force $pyinstallerDir }
if (Test-Path $releaseDir) { Remove-Item -Recurse -Force $releaseDir }
if (Test-Path $zipPath) { Remove-Item -Force $zipPath }

python -m PyInstaller --noconfirm --clean .\ExcelSyncManager.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $pyinstallerDir)) {
    throw "PyInstaller output folder not found: $pyinstallerDir"
}

New-Item -ItemType Directory -Force $releaseRoot | Out-Null
Copy-Item $pyinstallerDir $releaseDir -Recurse -Force
New-Item -ItemType Directory -Force (Join-Path $releaseDir "_runtime") | Out-Null
Copy-Item .\tasks.template.json (Join-Path $releaseDir "tasks.json") -Force
Copy-Item .\README_RELEASE.md (Join-Path $releaseDir "README_RELEASE.md") -Force

Compress-Archive -Path $releaseDir -DestinationPath $zipPath -Force

Write-Host "Build complete." -ForegroundColor Green
Write-Host "Raw dist folder: $pyinstallerDir"
Write-Host "Versioned release folder: $releaseDir"
Write-Host "Versioned release zip: $zipPath"
