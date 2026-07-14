#!/usr/bin/env bash
# Estimate workload and select engine (thin wrapper around estimate_workload.py)
# Usage: ./estimate-workload.sh -P 50 -C 30 [-T 5] [-I 0] [--type academic] [--toc]
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec python3 "$SCRIPT_DIR/estimate_workload.py" "$@"
