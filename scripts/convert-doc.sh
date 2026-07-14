#!/usr/bin/env bash
# Convert .doc to .docx (thin wrapper around convert_doc.py)
# Usage: ./convert-doc.sh -i template.doc [-o template.docx]
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec python3 "$SCRIPT_DIR/convert_doc.py" "$@"
