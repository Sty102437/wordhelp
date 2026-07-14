#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""wordhelp cross-platform installer.

Detects OS, installs Python dependencies, optionally builds minimax-docx backend.

Usage:
  python install.py            # full install
  python install.py --minimal  # python-docx only, skip minimax-docx
"""
import sys
import os
import platform
import subprocess
import argparse


def run(cmd, **kwargs):
    """Run a command and return (success, output)."""
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, **kwargs)
        return result.returncode == 0, result.stdout.strip() + result.stderr.strip()
    except FileNotFoundError:
        return False, "not found"


def check_python_docx():
    try:
        import docx
        return True, docx.__version__
    except ImportError:
        return False, None


def install_python_docx():
    print("  Installing python-docx...")
    success, output = run([sys.executable, "-m", "pip", "install", "python-docx"])
    return success


def check_dotnet():
    success, version = run(["dotnet", "--version"])
    if not success:
        return False, None
    try:
        major = int(version.split(".")[0])
        return (major >= 8), version
    except (ValueError, IndexError):
        return False, version


def install_pywin32():
    """Install pywin32 (Windows only)."""
    print("  Installing pywin32...")
    success, output = run([sys.executable, "-m", "pip", "install", "pywin32"])
    return success


def check_converter():
    """Check .doc converter availability."""
    system = platform.system()
    found = []

    if system == "Windows":
        try:
            import win32com.client
            app = win32com.client.Dispatch("KWPS.Application")
            app.Quit()
            found.append("WPS COM")
        except Exception:
            pass

    for cmd in ["soffice", "libreoffice"]:
        success, _ = run([cmd, "--version"])
        if success:
            found.append("LibreOffice")
            break

    return found


def find_minimax_skill():
    """Find minimax-docx skill directory."""
    home = os.path.expanduser("~")

    env_path = os.environ.get("WORDHELP_MINIMAX_SKILL")
    if env_path and os.path.isdir(env_path):
        return env_path

    candidates = [
        os.path.join(home, ".codex", "skills", "minimax-docx"),
        os.path.join(home, ".workbuddy", "skills", "minimax-docx"),
        os.path.join(home, ".trae", "skills", "minimax-docx"),
    ]

    for path in candidates:
        if os.path.isdir(path):
            return path

    return None


def build_minimax(skill_path):
    """Build minimax-docx backend."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    build_script = os.path.join(script_dir, "build_minimax.py")
    success, output = run([sys.executable, build_script, "--skill-path", skill_path])
    return success, output


def main():
    parser = argparse.ArgumentParser(description="wordhelp cross-platform installer")
    parser.add_argument("--minimal", action="store_true", help="python-docx only, skip minimax-docx")
    args = parser.parse_args()

    system = platform.system()
    print(f"\n=== wordhelp Installer ({system}) ===\n")

    ok = 0
    fail = 0

    # 1. Python + python-docx
    print("[1/4] Python + python-docx...", end=" ")
    success, version = check_python_docx()
    if success:
        print(f"OK ({version})")
        ok += 1
    else:
        print("not found")
        if install_python_docx():
            success, version = check_python_docx()
            if success:
                print(f"  Installed OK ({version})")
                ok += 1
            else:
                print("  FAILED — install manually: pip install python-docx")
                fail += 1
        else:
            print("  FAILED — install manually: pip install python-docx")
            fail += 1

    # 2. pywin32 (Windows only)
    if system == "Windows":
        print("[2/4] pywin32 (for .doc conversion)...", end=" ")
        try:
            import win32com.client
            print("OK")
            ok += 1
        except ImportError:
            print("not found")
            if install_pywin32():
                print("  Installed OK")
                ok += 1
            else:
                print("  FAILED — install manually: pip install pywin32")
                fail += 1
    else:
        print("[2/4] pywin32... SKIPPED (not needed on macOS/Linux)")
        ok += 1

    # 3. .NET SDK
    print("[3/4] .NET SDK 8.0+...", end=" ")
    success, version = check_dotnet()
    if success:
        print(f"OK ({version})")
        ok += 1
    elif version:
        print(f"FAILED (need >= 8.0, got {version})")
        fail += 1
    else:
        print("NOT FOUND")
        print("  Install: https://dotnet.microsoft.com/download")
        print("  Or: winget install Microsoft.DotNet.SDK.8 (Windows)")
        print("  Or: brew install --cask dotnet-sdk (macOS)")
        fail += 1

    # 4. minimax-docx backend
    if not args.minimal:
        print("[4/4] minimax-docx backend...", end=" ")
        skill_path = find_minimax_skill()
        if skill_path:
            print(f"found at {skill_path}")
            print("  Building...")
            success, output = build_minimax(skill_path)
            if success:
                print("  Build OK")
                ok += 1
            else:
                print(f"  Build FAILED: {output[:200]}")
                fail += 1
        else:
            print("NOT FOUND")
            print("  Set WORDHELP_MINIMAX_SKILL env var or install minimax-docx skill")
            print("  Or run with --minimal to skip (python-docx only mode)")
    else:
        print("[4/4] minimax-docx... SKIPPED (--minimal)")

    # Converter check
    print("\n--- Document Converter ---")
    converters = check_converter()
    if converters:
        print(f"  Available: {', '.join(converters)}")
    else:
        print("  WARNING: No .doc converter found.")
        if system == "Windows":
            print("  Install WPS Office or Microsoft Word for .doc conversion.")
        else:
            print("  Install LibreOffice: https://www.libreoffice.org/download/")

    print(f"\n=== Done ({ok} OK, {fail} failed) ===")
    print("Run smoke test: python scripts/smoke_test.py")

    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
