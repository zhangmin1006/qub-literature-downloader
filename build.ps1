# build.ps1 — Build LiteratureDownloader.exe
# Usage:  powershell -ExecutionPolicy Bypass -File build.ps1
# Output: dist\LiteratureDownloader.exe

Set-Location $PSScriptRoot

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Literature Auto-Downloader — EXE Builder" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
$py = (Get-Command python -ErrorAction SilentlyContinue)?.Source
if (-not $py) { Write-Error "Python not found in PATH"; exit 1 }
Write-Host "Python: $py"

# Ensure PyInstaller is available
$pi = python -c "import PyInstaller; print(PyInstaller.__version__)" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
    python -m pip install pyinstaller --quiet
}
Write-Host "PyInstaller: $pi"
Write-Host ""

# Clean previous build artefacts
if (Test-Path "build") { Remove-Item "build" -Recurse -Force }
if (Test-Path "dist\LiteratureDownloader.exe") {
    Remove-Item "dist\LiteratureDownloader.exe" -Force
}

# Run PyInstaller
Write-Host "Building EXE (this takes ~60 seconds)..." -ForegroundColor Yellow
$start = Get-Date
python -m PyInstaller LiteratureDownloader.spec --noconfirm
$elapsed = [int]((Get-Date) - $start).TotalSeconds

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Error "Build FAILED. Check output above for details."
    exit 1
}

$exe = "dist\LiteratureDownloader.exe"
if (-not (Test-Path $exe)) {
    Write-Error "EXE not found at $exe after build"
    exit 1
}

$sizeMB = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "  BUILD SUCCEEDED in ${elapsed}s" -ForegroundColor Green
Write-Host "  $((Resolve-Path $exe).Path)" -ForegroundColor Green
Write-Host "  Size: ${sizeMB} MB" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Double-click the EXE to launch the app." -ForegroundColor White
Write-Host "Your .lit_web_config.json will be saved next to the EXE." -ForegroundColor Gray
Write-Host ""
