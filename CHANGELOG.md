# Changelog

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
