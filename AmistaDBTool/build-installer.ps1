# build-installer.ps1
# Automated build script for AmistaDBTool installer
#
# Usage:
#   .\build-installer.ps1              - Full build (publish + installer)
#   .\build-installer.ps1 -SkipPublish - Only compile installer (assumes already published)
#

param(
    [switch]$SkipPublish,
    [string]$InnoSetupPath = "C:\Users\LAR00047\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
)

$ErrorActionPreference = "Stop"
$ProjectDir = $PSScriptRoot
$ProjectFile = Join-Path $ProjectDir "AmistaDBTool.csproj"
$PublishDir = Join-Path $ProjectDir "publish"
$SetupScript = Join-Path $ProjectDir "setup.iss"
$OutputDir = Join-Path $ProjectDir "Output"

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "AmistaDBTool Installer Build Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Validate Inno Setup
if (-not (Test-Path $InnoSetupPath)) {
    Write-Host "ERROR: Inno Setup not found at: $InnoSetupPath" -ForegroundColor Red
    Write-Host "Install Inno Setup 6 from https://jrsoftware.org/isinfo.php" -ForegroundColor Yellow
    Write-Host "Or specify path: .\build-installer.ps1 -InnoSetupPath 'C:\path\to\ISCC.exe'" -ForegroundColor Yellow
    exit 1
}

if (-not $SkipPublish) {
    # Publish application
    Write-Host "[1/2] Publishing application..." -ForegroundColor Green
    if (Test-Path $PublishDir) { Remove-Item $PublishDir -Recurse -Force }

    dotnet publish $ProjectFile -c Release -r win-x64 -o $PublishDir --no-self-contained
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Publish failed!" -ForegroundColor Red
        exit 1
    }
    Write-Host "      Published to: $PublishDir" -ForegroundColor Gray
    Write-Host ""
}
else {
    Write-Host "[1/2] Skipping publish (using existing files)..." -ForegroundColor Yellow
    Write-Host ""
}

# Compile Inno Setup installer
Write-Host "[2/2] Compiling Inno Setup installer..." -ForegroundColor Green

if (-not (Test-Path $SetupScript)) {
    Write-Host "ERROR: setup.iss not found at: $SetupScript" -ForegroundColor Red
    exit 1
}

& $InnoSetupPath $SetupScript
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Inno Setup compilation failed!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "BUILD COMPLETED SUCCESSFULLY!" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installer: $OutputDir\AmistaDBToolSetup.exe" -ForegroundColor Cyan
Write-Host ""
