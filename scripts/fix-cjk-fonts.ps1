param([Parameter(Mandatory=$true)]$InputPath, $OutputPath, $CnFont="SimSun", $EnFont="Times New Roman")
if (-not $OutputPath) { $ext = [System.IO.Path]::GetExtension($InputPath); $OutputPath = $InputPath.Replace($ext, ".fixed$ext") }

$pyFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.py'
@"
import sys; sys.path.insert(0, r'D:\Lib\site-packages')
from docx import Document; from docx.oxml.ns import qn; from docx.oxml import OxmlElement
import re
cn_p = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf]')
ascii_p = re.compile(r'[a-zA-Z0-9]')
cn_font = sys.argv[3]; en_font = sys.argv[4]
doc = Document(sys.argv[1])
for para in doc.paragraphs:
    for run in para.runs:
        if not run.text.strip(): continue
        has_cn = bool(cn_p.search(run.text))
        has_en = bool(ascii_p.search(run.text))
        if has_cn:
            run.font.name = cn_font
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None: rFonts = OxmlElement('w:rFonts'); rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), cn_font)
        if has_en:
            run.font.name = en_font
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None: rFonts = OxmlElement('w:rFonts'); rPr.insert(0, rFonts)
            rFonts.set(qn('w:ascii'), en_font); rFonts.set(qn('w:hAnsi'), en_font)
doc.save(sys.argv[2])
print('Fixed: ' + sys.argv[2])
"@ | Set-Content $pyFile -Encoding UTF8

& D:\python.exe $pyFile $InputPath $OutputPath $CnFont $EnFont
Remove-Item $pyFile -Force -ErrorAction SilentlyContinue
