<#
	Build script for PZ WS to Server

	- Generates a unique build ID (timestamp + git short hash) so the EXE always changes
	- Writes a temporary version resource file to stamp FileVersion/ProductVersion strings
	- Invokes PyInstaller to create a one-file, windowed EXE with bundled data files
	- Usage: powershell -ExecutionPolicy Bypass -File build.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-GitShortHash {
		try {
				$hash = (git rev-parse --short HEAD 2>$null).Trim()
				if (-not $hash) { return 'nogit' }
				return $hash
		} catch { return 'nogit' }
}

# Compute build metadata
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$gitHash = Get-GitShortHash
$buildId = "$timestamp-$gitHash"

# Paths
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$name = 'PZ Dedi Serv and Mod Manager'
$dist = Join-Path $root 'dist'
$work = Join-Path $root 'build'
$entry = Join-Path $root 'pz_mod_scraper.py'

# Temp artifacts for stamping
$tempDir = Join-Path $env:TEMP 'pz-ws-to-server-build'
if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir | Out-Null }

# Create a build ID file in repo root (relative path avoids drive-letter colon issues)
$buildIdLocal = Join-Path $root '.build_id.txt'
Set-Content -Path $buildIdLocal -Value $buildId -Encoding UTF8 -NoNewline

$versionFile = Join-Path $tempDir 'version_info.txt'
$versionInfo = @"
# UTF-8
VSVersionInfo(
	ffi=FixedFileInfo(
		filevers=(1, 0, 0, 0),
		prodvers=(1, 0, 0, 0),
		mask=0x3f,
		flags=0x0,
		OS=0x40004,
		fileType=0x1,
		subtype=0x0,
		date=(0, 0)
		),
	kids=[
		StringFileInfo([
			StringTable(u'040904B0', [
				StringStruct(u'CompanyName', u''),
				StringStruct(u'FileDescription', u'Fiears PZ WS to Server'),
				StringStruct(u'FileVersion', u'1.0.0+$buildId'),
				StringStruct(u'InternalName', u'$name'),
				StringStruct(u'LegalCopyright', u''),
				StringStruct(u'OriginalFilename', u'$name.exe'),
				StringStruct(u'ProductName', u'PZ WS to Server'),
				StringStruct(u'ProductVersion', u'1.0.0+$buildId')
			])
		]),
		VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
	]
)
"@
Set-Content -Path $versionFile -Value $versionInfo -Encoding UTF8

# Resolve Python
$pyArgs = @()
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if ($pythonCmd) {
	$python = $pythonCmd.Path
} else {
	$pythonCmd = Get-Command py -ErrorAction SilentlyContinue
	if ($pythonCmd) {
		$python = $pythonCmd.Path
		$pyArgs += '-3'
	} else {
		throw 'Python was not found on PATH. Please install Python or add it to PATH.'
	}
}

# Ensure PyInstaller is importable
try {
		& $python @($pyArgs + @('-m','PyInstaller','--version')) | Out-Null
} catch {
		Write-Error 'PyInstaller is not installed in this Python environment. Install it (pip install pyinstaller) and retry.'
		exit 1
}

# Construct PyInstaller args
$piArgs = @(
		$pyArgs
) + @(
		'-m','PyInstaller',
		'--noconfirm',
		'--onefile',
		'--windowed',
		'--name', $name,
		'--distpath', $dist,
		'--workpath', $work,
		'--version-file', $versionFile,
	'--add-data', 'AboutInfo.txt:.',
	'--add-data', 'Descriptions.json:.',
	'--add-data', 'Settings.json:.',
	'--add-data', '.build_id.txt:.',
		$entry
)

Write-Host "Building '$name' with build ID $buildId" -ForegroundColor Cyan
& $python $piArgs
$exit = $LASTEXITCODE
if ($exit -ne 0) {
		Write-Error "PyInstaller failed with exit code $exit"
	# Clean up temporary build id file on failure as well
	if (Test-Path $buildIdLocal) { Remove-Item $buildIdLocal -Force -ErrorAction SilentlyContinue }
	exit $exit
}

# Report result
$exePath = Join-Path $dist ("$name.exe")
if (Test-Path $exePath) {
		$sizeMB = '{0:N2}' -f ((Get-Item $exePath).Length / 1MB)
		Write-Host "Built: $exePath ($sizeMB MB)" -ForegroundColor Green
		Write-Host "Build ID: $buildId" -ForegroundColor Yellow
	# Clean up build id file
	if (Test-Path $buildIdLocal) { Remove-Item $buildIdLocal -Force -ErrorAction SilentlyContinue }
		exit 0
} else {
		Write-Error 'Build completed but EXE not found.'
	if (Test-Path $buildIdLocal) { Remove-Item $buildIdLocal -Force -ErrorAction SilentlyContinue }
		exit 1
}

