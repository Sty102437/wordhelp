param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [string]$OutputPath,
    [string]$CnFont = "SimSun",
    [string]$EnFont = "Times New Roman",
    [string]$PythonPath = $(if ($env:WORDHELP_PYTHON) { $env:WORDHELP_PYTHON } else { "D:\python.exe" }),
    [string]$SitePackages = $(if ($env:WORDHELP_SITE_PACKAGES) { $env:WORDHELP_SITE_PACKAGES } else { "D:\Lib\site-packages" })
)
if (-not $OutputPath) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $dir = [System.IO.Path]::GetDirectoryName($InputPath)
    $ext = [System.IO.Path]::GetExtension($InputPath)
    $OutputPath = Join-Path $dir "$base.fixed$ext"
}
if (-not (Test-Path $InputPath)) { Write-Error "Not found: $InputPath"; exit 1 }

$pyFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.py'
@"
import sys
sys.path.insert(0, r'$SitePackages')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

cn_p = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf]')
ascii_p = re.compile(r'[a-zA-Z0-9]')
cn_font = sys.argv[3]
en_font = sys.argv[4]

def fix_run(run):
    if not run.text.strip():
        return
    has_cn = bool(cn_p.search(run.text))
    has_en = bool(ascii_p.search(run.text))
    if not has_cn and not has_en:
        return
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    if has_cn:
        rFonts.set(qn('w:eastAsia'), cn_font)
    if has_en:
        rFonts.set(qn('w:ascii'), en_font)
        rFonts.set(qn('w:hAnsi'), en_font)

def fix_paragraphs(paragraphs):
    for para in paragraphs:
        for run in para.runs:
            fix_run(run)

doc = Document(sys.argv[1])

# paragraphs
fix_paragraphs(doc.paragraphs)

# table cells
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            fix_paragraphs(cell.paragraphs)

# headers and footers
for section in doc.sections:
    fix_paragraphs(section.header.paragraphs)
    fix_paragraphs(section.footer.paragraphs)

doc.save(sys.argv[2])
print('Fixed: ' + sys.argv[2])
"@ | Set-Content $pyFile -Encoding UTF8

& $PythonPath $pyFile $InputPath $OutputPath $CnFont $EnFont
Remove-Item $pyFile -Force -ErrorAction SilentlyContinue
