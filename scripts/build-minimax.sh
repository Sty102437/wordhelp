#!/usr/bin/env bash
# Build minimax-docx backend (thin wrapper around build_minimax.py)
# Usage: ./build-minimax.sh [--skill-path PATH]
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec python3 "$SCRIPT_DIR/build_minimax.py" "$@"
