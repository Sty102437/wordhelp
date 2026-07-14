#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Convert .doc to .docx — cross-platform.

Windows: WPS Office COM (KWPS.Application) or Word COM
macOS/Linux: LibreOffice CLI (soffice --headless --convert-to docx)

Usage: python convert_doc.py --input template.doc [--output template.docx]
"""
import sys
import os
import platform
import subprocess
import argparse


def convert_windows_wps(input_path, output_path):
    """Convert using WPS COM on Windows."""
    try:
        import win32com.client
    except ImportError:
        print("Error: pywin32 not installed. Run: pip install pywin32")
        return False

    try:
        app = win32com.client.Dispatch('KWPS.Application')
        app.Visible = False
        doc = app.Documents.Open(os.path.abspath(input_path))
        doc.SaveAs2(os.path.abspath(output_path), 12)  # 12 = wdFormatXMLDocument
        doc.Close()
        app.Quit()
        return True
    except Exception as e:
        print(f"WPS COM failed: {e}")
        return False


def convert_windows_word(input_path, output_path):
    """Convert using Microsoft Word COM on Windows."""
    try:
        import win32com.client
    except ImportError:
        return False

    try:
        app = win32com.client.Dispatch('Word.Application')
        app.Visible = False
        doc = app.Documents.Open(os.path.abspath(input_path))
        doc.SaveAs2(os.path.abspath(output_path), 12)
        doc.Close()
        app.Quit()
        return True
    except Exception as e:
        print(f"Word COM failed: {e}")
        return False


def convert_libreoffice(input_path, output_path):
    """Convert using LibreOffice CLI (macOS/Linux/Windows)."""
    # Find soffice binary
    candidates = [
        'soffice',
        'libreoffice',
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
    ]

    soffice = None
    for c in candidates:
        try:
            subprocess.run([c, '--version'], capture_output=True, check=True)
            soffice = c
            break
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue

    if soffice is None:
        print("Error: LibreOffice not found.")
        print("  Install: https://www.libreoffice.org/download/")
        return False

    # LibreOffice converts to same directory, then we move
    out_dir = os.path.dirname(os.path.abspath(output_path)) or '.'
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    expected_output = os.path.join(out_dir, base_name + '.docx')

    try:
        subprocess.run(
            [soffice, '--headless', '--convert-to', 'docx',
             '--outdir', out_dir, input_path],
            capture_output=True, check=True
        )
        # Rename if needed
        if expected_output != output_path and os.path.exists(expected_output):
            os.rename(expected_output, output_path)
        return os.path.exists(output_path)
    except subprocess.CalledProcessError as e:
        print(f"LibreOffice conversion failed: {e}")
        return False


def convert(input_path, output_path, converter=None):
    """Convert .doc to .docx.

    Args:
        converter: 'wps', 'word', 'libreoffice', or None (auto-detect)
    """
    if not os.path.exists(input_path):
        print(f"Error: File not found: {input_path}")
        return False

    # Converter lookup table
    converters = {
        'wps': convert_windows_wps,
        'word': convert_windows_word,
        'libreoffice': convert_libreoffice,
    }

    # If user specified a converter, use it directly
    if converter and converter in converters:
        if converters[converter](input_path, output_path):
            print(f"Converted ({converter}): {output_path}")
            return True
        print(f"Error: {converter} conversion failed.")
        return False

    # Auto-detect: try in priority order
    system = platform.system()
    if system == 'Windows':
        order = ['wps', 'word', 'libreoffice']
    else:
        order = ['libreoffice']

    for name in order:
        if converters[name](input_path, output_path):
            print(f"Converted ({name}): {output_path}")
            return True

    print(f"Error: No .doc converter available. Tried: {', '.join(order)}")
    return False


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert .doc to .docx (cross-platform)")
    parser.add_argument("--input", "-i", required=True, help="Input .doc file")
    parser.add_argument("--output", "-o", help="Output .docx file (default: same name, .docx extension)")
    parser.add_argument("--converter", "-c", choices=["wps", "word", "libreoffice"],
                        help="Preferred converter (skip auto-detection). Set during skill init.")
    args = parser.parse_args()

    if not args.output:
        args.output = os.path.splitext(args.input)[0] + '.docx'

    success = convert(args.input, args.output, args.converter)
    sys.exit(0 if success else 1)
