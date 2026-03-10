#!/bin/bash
# Launch Microsoft Excel with CoolProp libraries

ADDIN_DIR="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized"

# Detect architecture: both checks must confirm Apple Silicon
uname_check=$(uname -v)
cpu_check=$(sysctl -n machdep.cpu.brand_string)

if echo "$uname_check" | grep -qi "ARM64" && echo "$cpu_check" | grep -qi "Apple"; then
    echo "Architecture: Apple Silicon (ARM)"
    ln -sf "$ADDIN_DIR/libCoolProp_arm_64.dylib" /tmp/libCoolProp.dylib
else
    echo "Architecture: x86"
    ln -sf "$ADDIN_DIR/libCoolProp_x86_64.dylib" /tmp/libCoolProp.dylib
    ln -sf "$ADDIN_DIR/libCoolProp_x86_32.dylib" /tmp/libCoolProp_32bit.dylib
fi

DYLD_LIBRARY_PATH="$ADDIN_DIR" /Applications/Microsoft\ Excel.app/Contents/MacOS/Microsoft\ Excel
