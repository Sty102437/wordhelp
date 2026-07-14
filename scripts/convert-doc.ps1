param([Parameter(Mandatory=$true)]$InputPath, $OutputPath)
if (-not $OutputPath) { $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, ".docx") }
if (-not (Test-Path $InputPath)) { Write-Error "Not found: $InputPath"; exit 1 }

# Write Python script to temp file to avoid path encoding issues
$pyFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.py'
@"
import sys, os
sys.path.insert(0, 'D:/Lib/site-packages/win32')
sys.path.insert(0, 'D:/Lib/site-packages')
sys.path.insert(0, 'D:/Lib/site-packages/win32/lib')
os.add_dll_directory(r'D:\Lib\site-packages\pywin32_system32')
import win32com.client, json
app = win32com.client.Dispatch('KWPS.Application')
app.Visible = False
doc = app.Documents.Open(sys.argv[1])
doc.SaveAs2(sys.argv[2], 12)
doc.Close()
app.Quit()
print('Converted: ' + sys.argv[2])
"@ | Set-Content $pyFile -Encoding UTF8

& D:\python.exe $pyFile $InputPath $OutputPath
Remove-Item $pyFile -Force -ErrorAction SilentlyContinue
