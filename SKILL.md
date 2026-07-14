---
name: wordhelp
description: >
  Dual-engine Word document processing. Light tasks → python-docx, heavy tasks → minimax-docx.
  Covers template analysis, smart section detection, auto engine routing, CJK font fixing,
  .doc conversion, and professional formatting. Replaces course-paper-builder.
triggers:
  - Word
  - docx
  - document
  - 文档
  - 报告
  - 合同
  - 公文
  - 排版
  - 模板生成
  - 填表格
  - 写报告
version: 1.0.2
license: MIT
---

# wordhelp

Dual-engine Word document processing. Routes light tasks to python-docx and heavy tasks to minimax-docx automatically.

## Quick Reference

| Task | Engine | Path |
|------|--------|------|
| Fill a few cells, replace text | python-docx | Phase 3 light |
| Generate full report/paper | minimax-docx | Phase 3 heavy |
| Apply template formatting | minimax-docx | Pipeline C |
| Fix Chinese/English fonts | scripts/fix-cjk-fonts.ps1 | Phase 4 |
| Convert .doc to .docx | scripts/convert-doc.ps1 | Phase 1 |
| Check environment | scripts/smoke-test.ps1 | Pre-flight |

## Workflow

### Phase 0: Pre-flight

**Always** ask before touching any document:

1. Format preferences (if not specified):
   - CN font: Song Ti / Hei Ti / Microsoft YaHei / other
   - EN font: Times New Roman / Calibri / other
   - Size: 12pt / 11pt / other
   - Line spacing: 1.5x / 1.15x / fixed / other
   - Margins: academic (left 3.17cm) / standard (2.54cm) / other

2. Cover page (if template has one):
   - "The template has a cover page. Fill it? (yes/no)"
   - If yes: ask for title, subtitle, author, date, etc.
   - If no: leave untouched

Default: Chinese Academic from `references/formatting-defaults.md`.

### Phase 1: Template analysis

Convert .doc first: `powershell scripts/convert-doc.ps1 -InputPath template.doc`

Scan .docx with python-docx:
```python
from docx import Document
doc = Document("template.docx")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"P{i} [{p.style.name}] {p.text[:80]}")
for i, t in enumerate(doc.tables):
    print(f"T{i}: {len(t.rows)}r x {len(t.columns)}c")
```

Zone classification:
- **Cover**: first 1-2 pages, centered, metadata fields → fill only
- **Scoring**: grading tables, score columns → **skip completely**
- **TOC**: clear for regeneration
- **Body**: everything else

### Phase 2: Engine routing

```powershell
powershell scripts/estimate-workload.ps1 -P <paras> -T <tables> -C <cells> -I <images> -Type <type>
```

| Signal | Weight | Favors |
|--------|--------|--------|
| Paragraphs | min(N/10, 5) | minimax |
| Table cells | min(N/5, 5) | minimax |
| Images | min(N/3, 5) | python-docx |
| Needs TOC | +2 | minimax |
| Academic/Gov type | +3 | minimax |

score > 8 → minimax-docx | score <= 8 → python-docx

Always tell the user which engine and why.

### Phase 3: Execute

#### python-docx (light tasks)

**Two-pass.** Insert content first. Format in a second backward pass.

Classification table for Pass 2:

| Pattern | Style strategy |
|---|---|
| `\d+(\.\d+)+` | `Heading 1` → override 宋体 16pt bold justify |
| `\d+\.\d+` | `Heading 2` → override 宋体 15pt bold justify |
| `\d+\.\d+\.\d+` | `Heading 3` → override 宋体 14pt not-bold justify |
| `摘要` / `Abstract` title | NO heading style → 宋体 16pt bold center |
| `关键词:` / `Keywords:` | NO style → "关键词:" bold, content 12pt |
| `参考文献` title | NO heading style → 宋体 16pt bold center |
| `\[\d+\]` | 宋体 10.5pt justify |
| Body | 宋体 12pt justify, 1.5x spacing, 2-char indent |

**Why Heading styles?** TOC only detects built-in `Heading 1/2/3`.

**Why NOT for Abstract/References?** Heading styles auto-number in TOC → wrong.

**Images**: `doc.add_picture()`. Table captions above (宋体 10.5pt bold center), figure captions below.

**Equations**: center content, right-align number via tab stop. Format: (chapter-seq).

#### minimax-docx (heavy tasks)

```powershell
# Ensure backend
if (-not (Test-Path "$env:TEMP\minimax-docx-build")) {
    powershell scripts/build-minimax.ps1
}

$cli = "$env:TEMP\minimax-docx-build\MiniMaxAIDocx.Cli"

# Create from scratch
dotnet run --project $cli --no-build -c Release -- create --type report --output out.docx --title "Title"

# Template edit
dotnet run --project $cli --no-build -c Release -- edit fill-placeholders --input in.docx --output out.docx --data '{"key":"value"}'

# Apply template format
dotnet run --project $cli --no-build -c Release -- apply-template --input src.docx --template tmpl.docx --output out.docx
```

### Phase 4: Post-processing

1. **Run-level CJK font fix** (always last):
   ```powershell
   powershell scripts/fix-cjk-fonts.ps1 -InputPath out.docx -CnFont "宋体" -EnFont "Times New Roman"
   ```

2. **TOC**: warn to F9 (python-docx) or built-in (minimax-docx).

3. **Verify**:
   ```python
   doc = Document("output.docx")
   print(f"Paras:{len(doc.paragraphs)} Tables:{len(doc.tables)}")
   ```

## Key Rules

1. **Scoring pages are immutable.** Detect by "评分/score/考核/scoring" + grading tables. Skip.
2. **Cover pages: ask, don't generate.** User provides all cover text.
3. **Engine choice is always explained.** Tell user why.
4. **Heading styles first, then override.** TOC depends on Heading 1/2/3.
5. **Abstract/References: NO Heading styles.** Prevents TOC auto-numbering.
6. **Two-pass formatting.** Content insertion ≠ formatting. Don't mix.
7. **CJK fix is ALWAYS last.** Run-level glyph correction after all formatting.
8. **.doc → .docx first.** Never process .doc directly.

## Integration

| Skill | Relation |
|---|---|
| experiment-report | Complementary: lab report visuals + Feishu pipeline |
| thesis-slim | Complementary: matplotlib figures + English polishing |
| pdf | Downstream: preview/validate wordhelp output |
| minimax-docx | Backend: must be installed in Codex marketplace |

## Prerequisites

- Python 3.10+ with python-docx
- .NET SDK 8.0+
- WPS Office (optional, for .doc conversion)
- minimax-docx: Codex/Trae marketplace → "Word 文档生成" by MiniMaxAI

Run `scripts/smoke-test.ps1` to verify. Run `scripts/install.ps1` for one-click setup.

## Copyright

Engine routing and template analysis partially inspired by WorkBuddy (Tencent/CodeBuddy). Dependencies python-docx and minimax-docx retain MIT licenses. SKILL.md and scripts are original MIT.
