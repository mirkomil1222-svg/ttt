#!/bin/bash
# Build script for creating Rash Manager v2 macOS application bundle
# Usage: ./build_macos.sh

set -euo pipefail

GREEN="\033[0;32m"
RED="\033[0;31m"
NC="\033[0m"

log() {
  printf "%b%s%b\n" "$GREEN" "$1" "$NC"
}

err() {
  printf "%b%s%b\n" "$RED" "$1" "$NC" >&2
}

if [[ "${OSTYPE}" != "darwin"* ]]; then
  err "This script must be run on macOS."
  exit 1
fi

log "[1/5] Checking Python installation..."
if ! command -v python3 >/dev/null 2>&1; then
  err "Python 3 is not installed. Install it via https://www.python.org/ or Homebrew."
  exit 1
fi
python3 --version

declare -r PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJECT_DIR"

log "[2/5] Upgrading pip..."
python3 -m pip install --upgrade pip

log "[3/5] Installing project dependencies..."
python3 -m pip install -r requirements.txt

log "[4/5] Ensuring PyInstaller is installed..."
python3 -m pip install pyinstaller

log "[5/5] Building macOS app bundle (this may take several minutes)..."
python3 -m PyInstaller build_exe.spec --clean --noconfirm

log "Build complete!"
log "App bundle location: dist/RashManager.app"
log "Executable: dist/RashManager.app/Contents/MacOS/RashManager"
log "Remember to copy titul_bubble_koordinatalar_2480x3508.xlsx and Titul.pdf next to the app bundle."
log "For PDF support, install Poppler via: brew install poppler"
