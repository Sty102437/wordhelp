---
name: wordhelp
description: >
  Dual-engine Word document processing. Automatically routes light tasks to python-docx
  and heavy tasks to minimax-docx. Handles template analysis, cover page detection,
  scoring page skip, CJK font fixing, and .doc-to-.docx conversion.
triggers:
  - Word
  - docx
  - document
  - 文档
  - 报告
  - 合同
  - 公文
  - 排版
  - 用模板生成文档
  - 填表格
  - 写报告
version: 1.2.1
license: MIT
agent_created: true
---

# wordhelp

Dual-engine Word document processing. python-docx for light tasks, minimax-docx for heavy tasks. Cross-platform (Windows / macOS / Linux).

## Installation

```bash
# Install Python dependencies
pip install -r requirements.txt

# Full install (checks env + builds minimax-docx)
python scripts/install.py

# Minimal install (python-docx only, skip minimax-docx)
python scripts/install.py --minimal

# Verify environment
python scripts/smoke_test.py
```

**Platform-specific for .doc conversion:**
- Windows: WPS Office or Microsoft Word (via COM) or LibreOffice
- macOS/Linux: LibreOffice (`brew install --cask libreoffice` or `apt install libreoffice`)

## Workflow

### Phase 0: Collect preferences

Ask the user (if not provided):

1. **Formatting**: CN font / size / line spacing / margins / page size. Default: Chinese Academic (see `references/formatting-defaults.md`).
2. **Cover page** (if template has one): ask whether to fill it and what content (title, author, date, etc.). Do not auto-generate.
3. **Document converter** (if `.doc` files are involved): which tool to use for .doc → .docx conversion?
   - Windows: **WPS Office** / **Microsoft Word** / **LibreOffice**
   - macOS/Linux: **LibreOffice**
   - If not specified, auto-detect in priority order (WPS → Word → LibreOffice on Windows; LibreOffice on macOS/Linux).
   - Record the choice and pass `--converter <wps|word|libreoffice>` to `convert_doc.py` in Phase 1.

### Phase 1: Template analysis

1. If `.doc`, convert first:
   ```bash
   python scripts/convert_doc.py --input template.doc --converter <wps|word|libreoffice>
   # Omit --converter to auto-detect
   # Or wrapper: bash scripts/convert-doc.sh -i template.doc  (macOS/Linux)
   # Or wrapper: powershell scripts/convert-doc.ps1 -InputPath template.doc  (Windows)
   ```
   Auto-detects: WPS COM (Windows) → Word COM (Windows) → LibreOffice (all platforms).

2. Analyze structure: `python scripts/analyze-template.py --input template.docx`
   - Identify: **cover page**, **scoring page** (skip completely), **TOC**, **body**.
   - Scoring page keywords: "评分", "score", "考核", "评分标准".

### Phase 2: Route to engine

```bash
python scripts/estimate_workload.py -P <paragraphs> -C <cells> [-T <tables>] [-I <images>] --type <academic|general|government> [--toc]
```

**Scoring:** paragraphs min(P/10,5) + cells min(C/5,5) + images min(I/3,5) + TOC:2 + academic/government:3.
- score > 8 → **minimax-docx** (heavy)
- score <= 8 → **python-docx** (light)

**Fallback:** If minimax-docx is unavailable, degrade to python-docx and inform the user.

### Phase 3: Execute

#### python-docx (light workload)

Two-pass approach: insert content first, format in a separate pass.

```bash
python scripts/format-content.py --input filled.docx --output output.docx --type academic
```

**Heading style mapping** (applied by format-content.py):

| Pattern | Type | Style | Font override |
|---------|------|-------|---------------|
| `\d+(\.\d+)+` or `第.+章` | Heading 1 | `Heading 1` | 宋体 16pt bold justify |
| `\d+\.\d+` | Heading 2 | `Heading 2` | 宋体 15pt bold justify |
| `\d+\.\d+\.\d+` | Heading 3 | `Heading 3` | 宋体 14pt not-bold justify |
| `摘要` / `Abstract` | Abstract title | NO heading style, `outlineLvl=0` | 宋体 16pt bold center |
| `参考文献` / `References` | References title | NO heading style, `outlineLvl=0` | 宋体 16pt bold center |
| `关键词:` / `Keywords:` | Keywords | none | 12pt, prefix bold |
| `\[\d+\]` | Reference item | none | 宋体 10.5pt justify |
| else | Body text | none | 宋体 12pt justify |

**Why Heading styles?** TOC only detects `Heading 1/2/3`. Without them, TOC is empty.
**Why override?** Word defaults for Heading 1 are blue SimHei — wrong for academic papers.
**Why `outlineLvl` for Abstract/References?** They appear in TOC without auto-numbering (avoids "1 摘要").

#### minimax-docx (heavy workload)

1. Build backend (first time): `python scripts/build_minimax.py`
2. Create: `dotnet run --project <temp>/minimax-docx-build/MiniMaxAIDocx.Cli --no-build -c Release -- create --type report --output output.docx --title "Title"`
3. Edit template: `dotnet run ... -- edit fill-placeholders --input template.docx --output output.docx --data '{"name":"John"}'`

### Phase 4: Post-processing

1. **CJK font fix** (always last):
   ```bash
   python scripts/fix_cjk_fonts.py --input output.docx --cn-font "宋体" --en-font "Times New Roman"
   ```
   - Sets `w:rFonts` eastAsia for CN, ascii/hAnsi for EN per run.
   - Processes paragraphs, table cells, headers, and footers.
2. **Table/figure captions**: above table (宋体 10.5pt bold center), below figure (宋体 10.5pt center).
3. **TOC**: format-content.py sets `w:updateFields=true` — Word auto-updates TOC on open.
4. **Verify**: `python scripts/verify-output.py --input output.docx`

## Key rules

1. **Scoring pages are immutable** — never modify them.
2. **Cover pages are filled, not generated** — ask the user.
3. **Engine selection is transparent** — tell the user which engine and why.
4. **Apply built-in Heading style FIRST, then override** — TOC needs it.
5. **Abstract/References use `outlineLvl`, not Heading styles** — avoids TOC auto-numbering.
6. **CJK font fix is the final step** — after all other formatting.
7. **Two-pass formatting** — insert content first, format second.
8. **`.doc` files are converted first** — never process directly. Ask user for converter preference in Phase 0.

## Prerequisites

- Python 3.10+ with python-docx (`pip install -r requirements.txt`)
- .NET SDK 8.0+ (for minimax-docx, optional)
- WPS Office / Microsoft Word (Windows, for .doc conversion, optional) or LibreOffice (all platforms, optional)
- minimax-docx skill (optional, for heavy tasks)

Environment variable (optional):
- `WORDHELP_MINIMAX_SKILL` — minimax-docx skill directory

Run `python scripts/smoke_test.py` to verify the environment.

## Script Reference

All scripts have Python (primary) and shell wrapper (convenience) versions:

| Function | Python (cross-platform) | Windows wrapper | Unix wrapper |
|----------|------------------------|-----------------|---------------|
| Smoke test | `smoke_test.py` | `smoke-test.ps1` | `smoke-test.sh` |
| Convert .doc | `convert_doc.py` | `convert-doc.ps1` | `convert-doc.sh` |
| Fix fonts | `fix_cjk_fonts.py` | `fix-cjk-fonts.ps1` | `fix-cjk-fonts.sh` |
| Estimate workload | `estimate_workload.py` | `estimate-workload.ps1` | `estimate-workload.sh` |
| Build minimax | `build_minimax.py` | `build-minimax.ps1` | `build-minimax.sh` |
| Analyze template | `analyze-template.py` | — | — |
| Format content | `format-content.py` | — | — |
| Verify output | `verify-output.py` | — | — |
| Install | `install.py` | `install.ps1` | — |

## Copyright

Engine routing logic and template analysis were partially inspired by WorkBuddy (Tencent/CodeBuddy) built-in skills. Underlying engines python-docx and minimax-docx retain their original MIT licenses. SKILL.md and scripts are original MIT.
