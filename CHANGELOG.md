# Changelog

## v1.2.1 (2026-07-14)

- **convert_doc.py**: Added `--converter` flag (`wps` / `word` / `libreoffice`). Skip auto-detection when user specifies a preferred converter.
- **SKILL.md**: Phase 0 now collects converter preference alongside formatting and cover info. Agent passes `--converter` to `convert_doc.py` in Phase 1.

## v1.2.0 (2026-07-14)

### Cross-Platform Support
- **All core logic moved to Python.** PowerShell (.ps1) and bash (.sh) scripts are now thin wrappers around cross-platform Python scripts (.py).
- **.doc conversion** now auto-detects: WPS COM (Windows) → Word COM (Windows) → LibreOffice CLI (all platforms).
- **Cross-platform installer**: `python scripts/install.py` replaces PowerShell-only install. Detects OS, installs dependencies, auto-finds minimax-docx skill.
- Added `requirements.txt` for Python dependency management.
- Added bash wrappers: `convert-doc.sh`, `fix-cjk-fonts.sh`, `estimate-workload.sh`, `smoke-test.sh`, `build-minimax.sh`.

### New Python Scripts (Primary Entry Points)
- `scripts/convert_doc.py` — cross-platform .doc to .docx conversion
- `scripts/fix_cjk_fonts.py` — CJK/EN font fix (standalone, no PowerShell needed)
- `scripts/estimate_workload.py` — workload estimation
- `scripts/smoke_test.py` — cross-platform environment check
- `scripts/build_minimax.py` — minimax-docx backend build

### Bug Fixes (from v1.1.0)
- **fix-cjk-fonts**: Fixed font overwrite bug where mixed CJK/EN runs lost their Chinese font setting. Now uses `w:rFonts` attributes exclusively.
- **fix-cjk-fonts**: Added table cell and header/footer paragraph processing.
- **estimate-workload**: `-P` and `-C` parameters now mandatory with validation.

### Improvements (from v1.1.0)
- All hardcoded paths removed from Python scripts. Uses standard import paths.
- SKILL.md: Added installation section, script reference table, cross-platform conversion documentation.
- SKILL.md: TOC update via `w:updateFields=true`. Added `outlineLvl` for Abstract/References. Added fallback path.
- README / README_zh: Rewritten with cross-platform install instructions.

## v1.1.0 (2026-07-14)

- Fixed fix-cjk-fonts.ps1 font overwrite bug (w:rFonts approach)
- All PowerShell scripts path-parameterized (WORDHELP_PYTHON env vars)
- SKILL.md slimmed from ~267 to ~120 lines, inline code extracted to .py files
- estimate-workload.ps1: P and C mandatory
- Added analyze-template.py, format-content.py, verify-output.py, test-fontfix.py

## v1.0.1 (2026-07-14)

- Add one-click install script (`scripts/install.ps1`) with auto-dependency resolution
- Support `-Minimal` flag for python-docx-only deployment
- Update README with quick install instructions

## v1.0.0 (2026-07-14)

- Initial release.
- Dual-engine architecture: python-docx (light tasks) + minimax-docx (heavy tasks).
- Automatic workload estimation and engine routing.
- Smart template analysis with cover page detection and scoring page skip.
- Pre-flight formatting preference collection.
- CJK font fix (Song Ti / Times New Roman dual-pass).
- .doc to .docx conversion via WPS COM.
