#!/usr/bin/env bash
# Fix CJK/EN fonts in .docx (thin wrapper around fix_cjk_fonts.py)
# Usage: ./fix-cjk-fonts.sh -i output.docx [-o fixed.docx] [--cn-font 宋体] [--en-font "Times New Roman"]
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec python3 "$SCRIPT_DIR/fix_cjk_fonts.py" "$@"
