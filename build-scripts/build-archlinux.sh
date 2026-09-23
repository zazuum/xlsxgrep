#!/usr/bin/env bash
# Build an xlsxgrep pacman package locally on Arch Linux (uses PyInstaller + makepkg).
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_DIR"

echo "==> [1/6] Checking system dependencies..."
REQUIRED_PACMAN_PACKAGES=(python python-pip base-devel)
MISSING_PACMAN_PACKAGES=()
for pkg in "${REQUIRED_PACMAN_PACKAGES[@]}"; do
    if ! pacman -Qi "$pkg" >/dev/null 2>&1; then
        MISSING_PACMAN_PACKAGES+=("$pkg")
    fi
done

if [ "${#MISSING_PACMAN_PACKAGES[@]}" -gt 0 ]; then
    if ! command -v pacman >/dev/null 2>&1; then
        echo "Missing packages: ${MISSING_PACMAN_PACKAGES[*]}"
        echo "pacman not found; install these packages manually for your distro and re-run."
        exit 1
    fi
    echo "==> Installing missing packages: ${MISSING_PACMAN_PACKAGES[*]}"
    sudo pacman -Sy --needed --noconfirm "${MISSING_PACMAN_PACKAGES[@]}"
else
    echo "==> All required system packages are already installed."
fi

VENV_DIR=".build-venv"

echo "==> [2/6] Creating Python virtual environment..."
python -m venv "$VENV_DIR"
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
deactivate

echo "==> [5/6] Assembling package tree for version ${VERSION}..."
PKG_ROOT="$PROJECT_DIR/package-arch"
rm -rf "$PKG_ROOT"
mkdir -p "$PKG_ROOT/usr/bin" "$PKG_ROOT/usr/share/man/man1" "$PKG_ROOT/usr/share/doc/xlsxgrep"

cp dist/xlsxgrep "$PKG_ROOT/usr/bin/xlsxgrep"
gzip -kc docs/xlsxgrep.1 > "$PKG_ROOT/usr/share/man/man1/xlsxgrep.1.gz"
cp LICENSE "$PKG_ROOT/usr/share/doc/xlsxgrep/LICENSE"

echo "==> [6/6] Writing PKGBUILD and running makepkg..."
BUILD_DIR="$PROJECT_DIR/arch-build"
rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR"

cat > "$BUILD_DIR/PKGBUILD" <<EOF
# Maintainer: Ivan Cvitic <cviticivan@gmail.com>
pkgname=xlsxgrep
pkgver=${VERSION}
pkgrel=1
pkgdesc="CLI tool to search text in spreadsheet files"
arch=('x86_64')
url="https://github.com/zazuum/xlsxgrep"
license=('MIT')
depends=('glibc')
options=('!strip' '!debug')

package() {
    cp -a "${PKG_ROOT}/usr" "\${pkgdir}/"
}
EOF

if [ "$(id -u)" -eq 0 ]; then
    if ! id -u builduser >/dev/null 2>&1; then
        useradd -m builduser
    fi
    chown -R builduser:builduser "$BUILD_DIR" "$PKG_ROOT"
    su -s /bin/bash builduser -c "cd '$BUILD_DIR' && makepkg -f --noconfirm"
else
    (cd "$BUILD_DIR" && makepkg -f --noconfirm)
fi

mkdir -p dist
cp "$BUILD_DIR"/xlsxgrep-*.pkg.tar.* dist/

echo "==> Done."
echo "  Package: dist/xlsxgrep-${VERSION}-1-x86_64.pkg.tar.zst"
echo "  Install with: sudo pacman -U ./dist/xlsxgrep-${VERSION}-1-x86_64.pkg.tar.zst"
