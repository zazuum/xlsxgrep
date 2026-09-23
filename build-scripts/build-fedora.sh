#!/usr/bin/env bash
# Build an xlsxgrep .rpm package locally on Fedora (mirrors .github/workflows/release.yml).
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")/.."

echo "==> [1/6] Checking system dependencies..."
REQUIRED_DNF_PACKAGES=(python3 python3-pip rpm-build findutils)
MISSING_DNF_PACKAGES=()
for pkg in "${REQUIRED_DNF_PACKAGES[@]}"; do
    if ! rpm -q "$pkg" >/dev/null 2>&1; then
        MISSING_DNF_PACKAGES+=("$pkg")
    fi
done

if [ "${#MISSING_DNF_PACKAGES[@]}" -gt 0 ]; then
    if ! command -v dnf >/dev/null 2>&1; then
        echo "Missing packages: ${MISSING_DNF_PACKAGES[*]}"
        echo "dnf not found; install these packages manually for your distro and re-run."
        exit 1
    fi
    echo "==> Installing missing packages: ${MISSING_DNF_PACKAGES[*]}"
    sudo dnf install -y "${MISSING_DNF_PACKAGES[@]}"
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

echo "==> [5/6] Assembling .rpm package tree for version ${VERSION}..."
rm -rf rpmbuild package-rpm
mkdir -p rpmbuild/{BUILD,BUILDROOT,RPMS,SOURCES,SPECS,SRPMS}
mkdir -p package-rpm/usr/bin package-rpm/usr/share/man/man1 package-rpm/usr/share/doc/xlsxgrep

cp dist/xlsxgrep package-rpm/usr/bin/xlsxgrep
gzip -kc docs/xlsxgrep.1 > package-rpm/usr/share/man/man1/xlsxgrep.1.gz
cp LICENSE package-rpm/usr/share/doc/xlsxgrep/LICENSE

tar -C package-rpm -czf "rpmbuild/SOURCES/xlsxgrep-${VERSION}.tar.gz" .

cat > rpmbuild/SPECS/xlsxgrep.spec <<EOF
Name: xlsxgrep
Version: ${VERSION}
Release: 1%{?dist}
Summary: CLI tool to search text in spreadsheet files
License: MIT
BuildArch: x86_64
Source0: %{name}-%{version}.tar.gz

%global debug_package %{nil}
%undefine _missing_build_ids_terminate_build
%define __debug_install_post %{nil}

%description
CLI tool to search text in XLSX, XLS, XLSM, CSV, TSV and ODS files.

%prep
%setup -q -c -T
tar -xzf %{SOURCE0}

%install
mkdir -p %{buildroot}
cp -a . %{buildroot}/

%files
/usr
%license /usr/share/doc/xlsxgrep/LICENSE

%changelog
* $(date "+%a %b %d %Y") Ivan Cvitic <cviticivan@gmail.com> - ${VERSION}-1
- Local Fedora package build
EOF

deactivate

echo "==> [6/6] Building .rpm package..."
rpmbuild \
    --define "_topdir $(pwd)/rpmbuild" \
    --define "debug_package %{nil}" \
    -bb rpmbuild/SPECS/xlsxgrep.spec

mkdir -p dist
cp rpmbuild/RPMS/x86_64/*.rpm "dist/"

echo "==> Done."
echo "  Package: dist/xlsxgrep-${VERSION}-1.fc*.x86_64.rpm"
echo "  Install with: sudo dnf install ./dist/xlsxgrep-${VERSION}-1.fc*.x86_64.rpm"
