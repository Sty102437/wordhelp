#!/usr/bin/env bash
# Environment smoke test (thin wrapper around smoke_test.py)
# Usage: ./smoke-test.sh
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec python3 "$SCRIPT_DIR/smoke_test.py" "$@"
