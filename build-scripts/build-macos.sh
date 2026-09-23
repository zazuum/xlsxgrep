#!/usr/bin/env bash
# Build xlsxgrep standalone binary, .tar.gz and .pkg locally on macOS (mirrors .github/workflows/release.yml).
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")/.."

echo "==> [1/6] Checking dependencies..."
MISSING=()
command -v python3 >/dev/null 2>&1 || MISSING+=("python3")
xcode-select -p >/dev/null 2>&1 || MISSING+=("Xcode Command Line Tools")
command -v pkgbuild >/dev/null 2>&1 || MISSING+=("pkgbuild")

if [ "${#MISSING[@]}" -gt 0 ]; then
    echo "The following dependencies are missing: ${MISSING[*]}"
    if [[ " ${MISSING[*]} " == *"Xcode Command Line Tools"* ]]; then
        echo "Installing Xcode Command Line Tools (provides pkgbuild)..."
        xcode-select --install
        echo "Re-run this script after the Command Line Tools installation finishes."
    fi
    if [[ " ${MISSING[*]} " == *"python3"* ]]; then
        echo "Install Python 3 first, e.g.: brew install python@3.12"
    fi
    exit 1
fi
echo "==> All required dependencies are present."

VENV_DIR=".build-venv"

echo "==> [2/6] Creating Python virtual environment..."
python3 -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"

echo "==> [3/6] Installing xlsxgrep and PyInstaller..."
python -m pip install --upgrade pip
python -m pip install . pyinstaller

echo "==> [4/6] Running PyInstaller..."
pyinstaller --noconfirm --clean --onefile --name xlsxgrep \
    --collect-all pyexcel \
    --collect-all pyexcel_io \
    --hidden-import pyexcel_io.writers \
    --collect-all pyexcel_xls \
    --collect-all pyexcel_xlsx \
    --collect-all pyexcel_odsr \
    --collect-all openpyxl \
    --collect-all xlrd \
    xlsxgrep/xlsxgrep.py

VERSION=$(python -c "from xlsxgrep import __version__; print(__version__)")
deactivate

echo "==> [5/6] Packaging portable tarball for version ${VERSION}..."
mkdir -p dist
tar -C dist -czf "dist/xlsxgrep-${VERSION}-macos-arm64.tar.gz" xlsxgrep

echo "==> [6/6] Building .pkg installer..."
rm -rf package-macos
mkdir -p package-macos/usr/local/bin
cp dist/xlsxgrep package-macos/usr/local/bin/xlsxgrep
chmod +x package-macos/usr/local/bin/xlsxgrep

pkgbuild \
    --root package-macos \
    --identifier org.zazuum.xlsxgrep \
    --version "${VERSION}" \
    --install-location / \
    "dist/xlsxgrep-${VERSION}-macos-arm64.pkg"

echo "==> Done."
echo "  Binary:  dist/xlsxgrep"
echo "  Tarball: dist/xlsxgrep-${VERSION}-macos-arm64.tar.gz"
echo "  PKG:     dist/xlsxgrep-${VERSION}-macos-arm64.pkg"
