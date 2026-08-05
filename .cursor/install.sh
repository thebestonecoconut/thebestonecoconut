#!/usr/bin/env bash
#
# Cloud Agent install script for thebestonecoconut.
#
# This repository ships small, standalone utilities for BarTender label
# files (.btw):
#   * btw_convert.py         - Python 3 + olefile: extract text from .btw files
#   * kopiuj_etykiety.ps1     - PowerShell: copy .btw labels matching an Excel/CSV list
#
# The utilities live on feature branches; `main` may only contain the README.
# This script therefore prepares BOTH toolchains and stays safe (idempotent)
# even when the product files are not present on the checked-out branch.
set -euo pipefail

echo "==> Setting up development environment for thebestonecoconut"

# ---------------------------------------------------------------------------
# 1. PowerShell 7 (runs the .ps1 label-copier; the script only uses
#    cross-platform .NET APIs, so it works on Linux).
# ---------------------------------------------------------------------------
if command -v pwsh >/dev/null 2>&1; then
  echo "==> PowerShell already installed: $(pwsh --version)"
else
  echo "==> Installing PowerShell 7 via Microsoft apt repository"
  # shellcheck disable=SC1091
  . /etc/os-release
  tmpdeb="$(mktemp --suffix=.deb)"
  curl -fsSL -o "$tmpdeb" \
    "https://packages.microsoft.com/config/ubuntu/${VERSION_ID}/packages-microsoft-prod.deb"
  sudo dpkg -i "$tmpdeb"
  rm -f "$tmpdeb"
  sudo apt-get update -qq
  sudo apt-get install -y -qq powershell
  echo "==> Installed $(pwsh --version)"
fi

# ---------------------------------------------------------------------------
# 2. Python virtual environment + dependencies for the .btw text converter.
#    Guard the requirements.txt reference: it only exists on the converter
#    branch, so fall back to the single known dependency otherwise.
# ---------------------------------------------------------------------------
# Ensure the venv/ensurepip tooling is available (missing on some base images).
if ! python3 -c "import ensurepip" >/dev/null 2>&1; then
  echo "==> Installing python3 venv tooling"
  pyver="$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')"
  sudo apt-get update -qq
  sudo apt-get install -y -qq "python${pyver}-venv" python3-pip
fi

echo "==> Setting up Python virtual environment (.venv)"
python3 -m venv .venv
# shellcheck disable=SC1091
. .venv/bin/activate
python -m pip install --upgrade pip >/dev/null

if [ -f requirements.txt ]; then
  echo "==> Installing Python dependencies from requirements.txt"
  python -m pip install -r requirements.txt
else
  echo "==> requirements.txt not present on this branch; installing olefile directly"
  python -m pip install "olefile>=0.47"
fi

echo "==> Environment ready:"
echo "      python : $(python --version)"
echo "      olefile: $(python -c 'import olefile; print(olefile.__version__)')"
echo "      pwsh   : $(pwsh --version)"
echo ""
echo "==> Activate the Python venv with: source .venv/bin/activate"
