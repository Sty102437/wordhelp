#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Build minimax-docx backend — cross-platform.

Copies the minimax-docx .NET project from the skill directory to a temp
location and builds it with dotnet CLI.

Usage: python build_minimax.py [--skill-path PATH]
"""
import sys
import os
import platform
import shutil
import subprocess
import tempfile
import argparse


def default_skill_path():
    """Detect default minimax-docx skill path based on platform."""
    home = os.path.expanduser("~")

    # Check environment variable first
    env_path = os.environ.get("WORDHELP_MINIMAX_SKILL")
    if env_path and os.path.isdir(env_path):
        return env_path

    # Common locations
    candidates = [
        os.path.join(home, ".codex", "skills", "minimax-docx"),
        os.path.join(home, ".workbuddy", "skills", "minimax-docx"),
        os.path.join(home, ".trae", "skills", "minimax-docx"),
    ]

    for path in candidates:
        if os.path.isdir(path):
            return path

    return None


def build(skill_path, build_dir=None):
    """Build minimax-docx backend."""
    dotnet_src = os.path.join(skill_path, "scripts", "dotnet")

    if not os.path.isdir(dotnet_src):
        print(f"Error: minimax-docx dotnet source not found: {dotnet_src}")
        return False

    if build_dir is None:
        build_dir = os.path.join(tempfile.gettempdir(), "minimax-docx-build")

    print(f"Building minimax-docx...")
    print(f"  Source: {dotnet_src}")
    print(f"  Output: {build_dir}")

    # Clean and copy
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    shutil.copytree(dotnet_src, build_dir)

    # Find .csproj
    cli_dir = os.path.join(build_dir, "MiniMaxAIDocx.Cli")
    csproj = os.path.join(cli_dir, "MiniMaxAIDocx.Cli.csproj")

    if not os.path.exists(csproj):
        print(f"Error: .csproj not found: {csproj}")
        return False

    # Restore + build
    print("  Restoring packages...")
    result = subprocess.run(
        ["dotnet", "restore", csproj, "--verbosity", "quiet"],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  Restore failed: {result.stderr}")
        return False

    print("  Building (Release)...")
    result = subprocess.run(
        ["dotnet", "build", csproj, "-c", "Release", "--no-restore", "--verbosity", "quiet"],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  Build FAILED: {result.stderr}")
        print("  Check .NET SDK and minimax-docx skill path.")
        return False

    print(f"  Done. Backend: {build_dir}")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build minimax-docx backend (cross-platform)")
    parser.add_argument("--skill-path", help="Path to minimax-docx skill directory")
    parser.add_argument("--build-dir", help="Build output directory (default: temp)")
    args = parser.parse_args()

    skill_path = args.skill_path or default_skill_path()

    if not skill_path:
        print("Error: minimax-docx skill not found.")
        print("  Set --skill-path or env WORDHELP_MINIMAX_SKILL")
        print("  Or run with --minimal to skip (python-docx only mode)")
        sys.exit(1)

    success = build(skill_path, args.build_dir)
    sys.exit(0 if success else 1)
