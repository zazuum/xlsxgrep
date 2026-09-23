#!/usr/bin/env bash
# Build an xlsxgrep .deb package locally on Debian/Ubuntu (mirrors .github/workflows/release.yml).
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")/.."

echo "==> [1/6] Checking system dependencies..."
REQUIRED_APT_PACKAGES=(python3 python3-venv python3-pip dpkg-dev)
MISSING_APT_PACKAGES=()
for pkg in "${REQUIRED_APT_PACKAGES[@]}"; do
    if ! dpkg -s "$pkg" >/dev/null 2>&1; then
        MISSING_APT_PACKAGES+=("$pkg")
    fi
done

if [ "${#MISSING_APT_PACKAGES[@]}" -gt 0 ]; then
    if ! command -v apt-get >/dev/null 2>&1; then
        echo "Missing packages: ${MISSING_APT_PACKAGES[*]}"
        echo "apt-get not found; install these packages manually for your distro and re-run."
        exit 1
    fi
    echo "==> Installing missing packages: ${MISSING_APT_PACKAGES[*]}"
    sudo apt-get update
    sudo apt-get install -y "${MISSING_APT_PACKAGES[@]}"
else
    echo "==> All required system packages are already installed."
fi

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
    --collect-all pyexcel_xls \
    --collect-all pyexcel_xlsx \
    --collect-all pyexcel_odsr \
    --collect-all openpyxl \
    --collect-all xlrd \
    xlsxgrep/xlsxgrep.py

VERSION=$(python -c "from xlsxgrep import __version__; print(__version__)")

echo "==> [5/6] Assembling .deb package tree for version ${VERSION}..."
rm -rf package-deb
mkdir -p package-deb/DEBIAN package-deb/usr/bin package-deb/usr/share/man/man1 package-deb/usr/share/doc/xlsxgrep
cp dist/xlsxgrep package-deb/usr/bin/xlsxgrep
gzip -kc docs/xlsxgrep.1 > package-deb/usr/share/man/man1/xlsxgrep.1.gz
cp LICENSE package-deb/usr/share/doc/xlsxgrep/copyright
printf 'Package: xlsxgrep\nVersion: %s\nArchitecture: amd64\nMaintainer: Ivan Cvitic <cviticivan@gmail.com>\nDescription: CLI tool to search text in spreadsheet files\n' "$VERSION" > package-deb/DEBIAN/control

deactivate

echo "==> [6/6] Building .deb package..."
mkdir -p dist
dpkg-deb --build package-deb "dist/xlsxgrep-${VERSION}-amd64.deb"

echo "==> Done."
echo "  Package: dist/xlsxgrep-${VERSION}-amd64.deb"
echo "  Install with: sudo apt install ./dist/xlsxgrep-${VERSION}-amd64.deb"
