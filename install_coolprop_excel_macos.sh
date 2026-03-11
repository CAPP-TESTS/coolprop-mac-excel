#!/bin/bash
# Install CoolProp Excel wrapper on macOS

set -e

BASEURL="https://raw.githubusercontent.com/CAPP-TESTS/coolprop-mac-excel/refs/heads/main"
ADDIN_DIR="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized"
DESKTOP="$HOME/Desktop"

download() {
    local filename="$1"
    local destination="$2"
    printf 'Downloading %s\n' "$filename"
    curl -fSL -o "$destination/$filename" "$BASEURL/$filename"
}

echo "Installing CoolProp 32/64bit for MacOS"

mkdir -p "$ADDIN_DIR"

download "CoolProp_RST.xlam"          "$ADDIN_DIR"
download "libCoolProp_arm_64.dylib"   "$ADDIN_DIR"
download "libCoolProp_x86_64.dylib"   "$ADDIN_DIR"
download "libCoolProp_x86_32.dylib"   "$ADDIN_DIR"
download "launch_excel_coolprop.sh" "$DESKTOP"

echo "Add-in location: $ADDIN_DIR"
