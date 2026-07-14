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
version: 1.0.0
license: MIT
---

# wordhelp

Dual-engine Word document processing toolkit. Automatically selects python-docx for light tasks and minimax-docx for heavy tasks.

## Workflow

### Phase 0: Collect preferences

Before touching any document, ask the user:

1. **Formatting preferences** (if not provided):
   - Font (CN): Song Ti / Microsoft YaHei / SimHei / other
   - Font size: 12pt (小四) / 11pt / other
   - Line spacing: 1.5x / 1.15x / other
   - Page margins: standard (2.54cm) / academic (left 3.17cm) / other
   - Page size: A4 / Letter

2. **Cover page** (if template has one):
   - Ask: "The template has a cover page. Would you like me to fill it?"
   - If yes: ask for title, subtitle, author, date, and any other fields visible on the cover
   - If no: leave the cover page untouched

If the user does not specify, use Chinese Academic defaults (see references/formatting-defaults.md).

### Phase 1: Template analysis

1. If input is `.doc`, convert to `.docx` first:
   ```powershell
   powershell scripts/convert-doc.ps1 -InputPath template.doc
   ```

2. Analyze the resulting `.docx` to identify sections:
   - **Cover page**: first 1-2 pages with title, author, date fields (minimal body text, centered layout)
   - **Scoring page**: contains tables with scoring criteria, grade columns, teacher signature — **skip completely, do not modify**
   - **Table of Contents**: if present, clear it for regeneration after body is complete
   - **Body**: everything after the above sections

   Use python-docx to read paragraph texts and table structures:
   ```python
   from docx import Document
   doc = Document("template.docx")
   for i, p in enumerate(doc.paragraphs):
       if p.text.strip():
           print(f"P{i} [{p.style.name}] {p.text[:80]}")
   for i, t in enumerate(doc.tables):
       print(f"T{i}: {len(t.rows)}r x {len(t.columns)}c, first cell: {t.cell(0,0).text[:50]}")
   ```

### Phase 2: Route to engine

Use `scripts/estimate-workload.ps1` to evaluate task complexity:

```powershell
powershell scripts/estimate-workload.ps1 -P <paragraph_count> -T <table_count> -C <total_cells> -I <image_count> -Type <document_type>
```

**Scoring rules:**
- Paragraphs: min(paragraphCount / 10, 5) points
- Table cells: min(cellCount / 5, 5) points
- Images: min(imageCount / 3, 5) points
- TOC needed: +2 points
- Academic/Government type: +3 points

**Decision:**
- score > 8 → **minimax-docx** (heavy: batch generation, complex layouts, strict formatting)
- score <= 8 → **python-docx** (light: cell fills, text replacement, simple edits)

Tell the user which engine was selected and why. Example:
> "Estimated workload: 12.5 points — this is a heavy task (50+ cells in tables, academic format). I'll use minimax-docx for professional output. OK to proceed?"

### Phase 3: Execute

#### When using python-docx (light workload)

**Content filling — two rules:**
1. Insert content first, format in a separate pass. Formatting during insertion is unreliable because existing paragraph styles bleed into new content.
2. Classify paragraphs by text pattern before applying styles:

| Pattern | Type | Style Strategy |
|---------|------|----------------|
| `\d+(\.\d+)+` or `第.+章` | Heading 1 | Apply `Heading 1` style, then override: 宋体 16pt bold justify |
| `\d+\.\d+` | Heading 2 | Apply `Heading 2` style, then override: 宋体 15pt bold justify |
| `\d+\.\d+\.\d+` | Heading 3 | Apply `Heading 3` style, then override: 宋体 14pt not-bold justify |
| `摘要` or `Abstract` title | Abstract title | Do NOT apply heading style — manually set 宋体 16pt bold center |
| `关键词:` or `Keywords:` | Keywords line | No style — font 12pt, "关键词:" bold, content regular |
| `参考文献` | References title | Do NOT apply heading style — manually set 宋体 16pt bold center |
| `\[\d+\]` | Reference item | No style — 宋体 10.5pt justify |
| Everything else | Body text | No heading style — set font directly |

**Why use Heading styles?** Word Table of Contents only detects built-in styles (`Heading 1/2/3`). Without them, TOC is empty even if visual formatting looks correct.

**Why override?** Word defaults for `Heading 1` are blue SimHei — not what academic papers need. Apply the style label, then immediately override font/size/bold.

**Why NOT for Abstract/References titles?** Heading styles auto-number in TOC (e.g., "1 Abstract", "2 References"). These section titles should appear in TOC but NOT with chapter numbers.

```python
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document("template.docx")

# --- Two-pass approach ---
# Pass 1: insert content
# (fill paragraphs, tables, etc.)

# Pass 2: format with heading style mapping
for para in doc.paragraphs:
    text = para.text.strip()
    if not text: continue

    if is_heading1(text):
        para.style = doc.styles['Heading 1']
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in para.runs:
            run.font.name = '宋体'
            run.font.size = Pt(16)
            run.font.bold = True
    elif is_heading2(text):
        para.style = doc.styles['Heading 2']
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in para.runs:
            run.font.name = '宋体'
            run.font.size = Pt(15)
            run.font.bold = True
    elif is_heading3(text):
        para.style = doc.styles['Heading 3']
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in para.runs:
            run.font.name = '宋体'
            run.font.size = Pt(14)
            run.bold = False  # explicitly not bold
    elif is_abstract_title(text):
        # NO heading style — avoid TOC auto-numbering
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in para.runs:
            run.font.name = '宋体'
            run.font.size = Pt(16)
            run.font.bold = True
    elif is_references_title(text):
        # NO heading style
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in para.runs:
            run.font.name = '宋体'
            run.font.size = Pt(16)
            run.font.bold = True
    elif is_body(text):
        for run in para.runs:
            run.font.name = '宋体'
            run.font.size = Pt(12)
            run.bold = False
```

**Equation numbering** (if present):
```python
# Center the equation, right-align the number via tab stop
# Format: (1-1), (2-3) where first digit = chapter, second = equation sequence
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
# Add a right-aligned tab stop at page width for equation numbers
para.paragraph_format.tab_stops.add_tab_stop(
    Cm(14.66), WD_TAB_ALIGNMENT.RIGHT
)
```

#### When using minimax-docx (heavy workload)

1. Ensure backend is built:
   ```powershell
   if (-not (Test-Path "$env:TEMP\minimax-docx-build")) {
       powershell scripts/build-minimax.ps1
   }
   ```

2. Create professional document:
   ```powershell
   $cli = "$env:TEMP\minimax-docx-build\MiniMaxAIDocx.Cli"
   dotnet run --project $cli --no-build --configuration Release -- create --type report --output output.docx --title "Document Title"
   ```

3. For template-based work, use Pipeline C (apply-template) or Pipeline B (edit):
   ```powershell
   dotnet run --project $cli --no-build --configuration Release -- edit fill-placeholders --input template.docx --output output.docx --data '{"name":"John"}'
   ```

### Phase 4: Post-processing

1. **Run-level CJK font fix** — always run after any engine output. This is separate from paragraph-level formatting: it scans each individual text run and sets CN characters to Song Ti and EN/number characters to Times New Roman at the glyph level. Prevents mixed-font paragraphs where Chinese and English inherit different font families:
   ```powershell
   powershell scripts/fix-cjk-fonts.ps1 -InputPath output.docx -CnFont "宋体" -EnFont "Times New Roman"
   ```
   Performed after all other formatting because:
   - Paragraph-level formatting sets intent (heading/body)
   - Run-level fix enforces consistency (every character gets the right font)
   - Must be the last formatting step to avoid being overwritten

2. **Table & figure captions** (if present):
   - Table captions: above the table, 宋体 10.5pt bold center ("表1-1 数据对比")
   - Figure captions: below the figure, 宋体 10.5pt center ("图1-3 系统架构图")

3. **TOC regeneration** (if template had one):
   - For python-docx: warn user to open in Word and press F9
   - For minimax-docx: built-in TOC field generation

4. **Verify output**:
   ```python
   from docx import Document
   doc = Document("output.docx")
   print(f"Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}")
   for p in doc.paragraphs[:5]:
       if p.text.strip():
           print(f"  [{p.style.name}] {p.text[:80]}")
   ```

## Key rules

1. **Scoring pages are immutable.** They contain evaluation criteria that must not be altered. Detect them by scanning for keywords like "评分", "score", "考核", "评分标准", "scoring standard", and tables with numeric grade columns.

2. **Cover pages are filled, not generated.** Ask the user what to put there. Do not auto-generate cover content.

3. **Engine selection is transparent.** Always tell the user which engine was chosen and why.

4. **Heading styles: apply built-in style FIRST, then override formatting.** Word TOC detects `Heading 1/2/3` styles. Skipping the built-in style makes the heading invisible to TOC.

5. **Abstract and References titles do NOT use Heading styles.** They would get unwanted auto-numbering in TOC (e.g., "1 摘要").

6. **CJK font fix is the final step.** Always run run-level CN/EN font correction after all other formatting. Run-level fix guarantees every glyph gets the right font regardless of paragraph-level settings.

7. **Two-pass formatting.** Insert content first, format in a separate backward iteration. Direct-format contamination from previous paragraphs (bold, centering) bleeds into newly inserted content.

8. **.doc files are converted first.** Never attempt to process .doc directly. Convert via WPS COM.

## Prerequisites

- Python 3.10+ with python-docx: `D:\python.exe -c "import docx"`
- .NET SDK 8.0+: `dotnet --version`
- WPS Office (for .doc conversion): `D:\python.exe -c "import win32com.client; app=win32com.client.Dispatch('KWPS.Application')"`
- minimax-docx skill installed at `~\.codex\skills\minimax-docx\`

Run `scripts/smoke-test.ps1` to verify the environment.

## Copyright

Engine routing logic and template analysis were partially inspired by WorkBuddy (Tencent/CodeBuddy) built-in skills. Underlying engines python-docx and minimax-docx retain their original MIT licenses. SKILL.md and scripts are original MIT.
