# Build script for "Fiears PZ WS to Server.exe" (one-file, GUI)
param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

if ($Clean) {
    if (Test-Path 'build') { Remove-Item -Recurse -Force 'build' }
    if (Test-Path 'dist') { Remove-Item -Recurse -Force 'dist' }
    if (Test-Path '*.spec') { Remove-Item -Force *.spec }
}

# Ensure Python is available
try {
    $pyVersion = & python --version
    Write-Host "Using $pyVersion"
} catch {
    Write-Error 'Python was not found in PATH. Please install Python 3.10+ and try again.'
    exit 1
}

# Ensure PyInstaller is installed and invokable via module
Write-Host 'Checking PyInstaller...'
$pyiOk = $false
try {
    & python -m PyInstaller --version | Out-Null
    $pyiOk = $true
} catch {
    $pyiOk = $false
}

if (-not $pyiOk) {
    Write-Host 'Installing PyInstaller...'
    & python -m pip install --upgrade pip
    & python -m pip install pyinstaller
}

# Build
$exeName = 'Fiears PZ WS to Server'
$script = 'pz_mod_scraper.py'

# Use --windowed to hide console for the Tkinter GUI and include AboutInfo.txt as bundled data
$addData = @()
if (Test-Path 'AboutInfo.txt') {
    # Format: SRC;DEST on Windows for --add-data
    $addData = @('--add-data', "AboutInfo.txt;.")
}
& python -m PyInstaller --onefile --windowed --name $exeName @addData $script

# Copy helpful files next to the EXE (optional at runtime)
if (Test-Path 'AboutInfo.txt') { Copy-Item 'AboutInfo.txt' -Destination 'dist' -Force }
if (Test-Path 'README.md') { Copy-Item 'README.md' -Destination 'dist' -Force }

Write-Host "Build complete. Output: dist/$exeName.exe"
