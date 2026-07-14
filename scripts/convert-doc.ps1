param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [string]$OutputPath,
    [string]$PythonPath = $(if ($env:WORDHELP_PYTHON) { $env:WORDHELP_PYTHON } else { "D:\python.exe" }),
    [string]$SitePackages = $(if ($env:WORDHELP_SITE_PACKAGES) { $env:WORDHELP_SITE_PACKAGES } else { "D:\Lib\site-packages" })
)
if (-not $OutputPath) { $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, ".docx") }
if (-not (Test-Path $InputPath)) { Write-Error "Not found: $InputPath"; exit 1 }

# Write Python script to temp file to avoid path encoding issues
$pyFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.py'
@"
import sys, os
sys.path.insert(0, r'$SitePackages\win32')
sys.path.insert(0, r'$SitePackages')
sys.path.insert(0, r'$SitePackages\win32\lib')
os.add_dll_directory(r'$SitePackages\pywin32_system32')
import win32com.client
app = win32com.client.Dispatch('KWPS.Application')
app.Visible = False
doc = app.Documents.Open(sys.argv[1])
doc.SaveAs2(sys.argv[2], 12)
doc.Close()
app.Quit()
print('Converted: ' + sys.argv[2])
"@ | Set-Content $pyFile -Encoding UTF8

& $PythonPath $pyFile $InputPath $OutputPath
Remove-Item $pyFile -Force -ErrorAction SilentlyContinue
