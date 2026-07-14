#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Environment smoke test — cross-platform.

Checks: Python + python-docx, .NET SDK, document converter availability.

Usage: python smoke_test.py
"""
import sys
import platform
import subprocess


def check_python_docx():
    """Check Python and python-docx."""
    try:
        import docx
        return True, f"OK ({docx.__version__})"
    except ImportError:
        return False, "FAILED — pip install python-docx"


def check_dotnet():
    """Check .NET SDK 8.0+."""
    try:
        result = subprocess.run(["dotnet", "--version"], capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            version = result.stdout.strip()
            major = int(version.split(".")[0])
            if major >= 8:
                return True, f"OK ({version})"
            else:
                return False, f"FAILED (need >= 8.0, got {version})"
        return False, "FAILED"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False, "NOT FOUND — install from https://dotnet.microsoft.com/download"


def check_converter():
    """Check .doc converter availability."""
    system = platform.system()
    found = []

    if system == "Windows":
        # Check WPS COM
        try:
            import win32com.client
            app = win32com.client.Dispatch("KWPS.Application")
            app.Quit()
            found.append("WPS COM")
        except Exception:
            pass

        # Check Word COM
        try:
            import win32com.client
            app = win32com.client.Dispatch("Word.Application")
            app.Quit()
            found.append("Word COM")
        except Exception:
            pass

    # Check LibreOffice (all platforms)
    for cmd in ["soffice", "libreoffice"]:
        try:
            subprocess.run([cmd, "--version"], capture_output=True, check=True, timeout=10)
            found.append("LibreOffice")
            break
        except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
            continue

    if found:
        return True, f"OK ({', '.join(found)})"
    else:
        return False, "WARN (optional — install WPS/Word/LibreOffice for .doc conversion)"


def main():
    print("=== wordhelp Smoke Test ===")
    ok = 0
    bad = 0
    warn = 0

    checks = [
        ("Python + python-docx", check_python_docx),
        (".NET SDK 8.0+", check_dotnet),
        ("Document converter", check_converter),
    ]

    for i, (name, check_fn) in enumerate(checks, 1):
        success, msg = check_fn()
        status = "OK" if success else ("WARN" if "WARN" in msg else "FAILED")
        color = "\033[92m" if success else ("\033[93m" if "WARN" in msg else "\033[91m")
        reset = "\033[0m"
        print(f"[{i}/{len(checks)}] {name}... {color}{msg}{reset}")
        if success:
            ok += 1
        elif "WARN" in msg:
            warn += 1
        else:
            bad += 1

    print()
    if bad == 0:
        print(f"All checks passed! ({ok} OK, {warn} warnings)")
    else:
        print(f"{bad} failure(s), {warn} warning(s)")

    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
