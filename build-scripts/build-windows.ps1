# Build xlsxgrep.exe, portable zip and installer locally on Windows (mirrors .github/workflows/release.yml).
$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

Write-Host "==> [1/6] Checking dependencies..."
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python is required. Install it first, e.g.: winget install Python.Python.3.12"
    exit 1
}
Write-Host "==> All required dependencies are present."

$VenvDir = ".build-venv"

Write-Host "==> [2/6] Creating Python virtual environment..."
python -m venv $VenvDir
& "$VenvDir\Scripts\Activate.ps1"

Write-Host "==> [3/6] Installing xlsxgrep and PyInstaller..."
python -m pip install --upgrade pip
python -m pip install . pyinstaller

Write-Host "==> [4/6] Running PyInstaller..."
pyinstaller --noconfirm --clean --onefile --name xlsxgrep `
    --collect-all pyexcel `
    --collect-all pyexcel_io `
    --collect-all pyexcel_xls `
    --collect-all pyexcel_xlsx `
    --collect-all pyexcel_odsr `
    --collect-all openpyxl `
    --collect-all xlrd `
    xlsxgrep/xlsxgrep.py

$Version = python -c "from xlsxgrep import __version__; print(__version__)"

Write-Host "==> [5/6] Packaging portable zip for version $Version..."
Compress-Archive -Path "dist/xlsxgrep.exe" -DestinationPath "dist/xlsxgrep-$Version-windows-amd64.zip" -Force

deactivate

Write-Host "==> [6/6] Building installer with Inno Setup..."
$Iscc = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "${env:ProgramFiles}\Inno Setup 6\ISCC.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($Iscc) {
    & $Iscc.FullName "/DAppVersion=$Version" "installer\windows\xlsxgrep.iss"
    Move-Item -Force "installer\windows\Output\xlsxgrep-setup.exe" "dist/xlsxgrep-$Version-windows-setup.exe"
} else {
    Write-Host "Inno Setup not found; skipping installer build."
    Write-Host "Install it with: choco install innosetup   (or download from https://jrsoftware.org/isinfo.php)"
}

Write-Host "==> Done."
Write-Host "  Binary:    dist/xlsxgrep.exe"
Write-Host "  Zip:       dist/xlsxgrep-$Version-windows-amd64.zip"
Write-Host "  Installer: dist/xlsxgrep-$Version-windows-setup.exe (if Inno Setup was found)"
