param(
    [string]$PythonPath = $(if ($env:WORDHELP_PYTHON) { $env:WORDHELP_PYTHON } else { "D:\python.exe" }),
    [string]$SitePackages = $(if ($env:WORDHELP_SITE_PACKAGES) { $env:WORDHELP_SITE_PACKAGES } else { "D:\Lib\site-packages" })
)
Write-Host "=== wordhelp Smoke Test ===" -ForegroundColor Cyan
$ok = 0; $bad = 0

Write-Host "[1/3] Python + python-docx..." -NoNewline
$py = & $PythonPath -c "import sys; sys.path.insert(0, r'$SitePackages'); import docx; print(docx.__version__)" 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK ($py)" -ForegroundColor Green; $ok++ }
else { Write-Host " FAILED" -ForegroundColor Red; $bad++ }

Write-Host "[2/3] .NET SDK..." -NoNewline
$dot = dotnet --version 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK ($dot)" -ForegroundColor Green; $ok++ }
else { Write-Host " FAILED" -ForegroundColor Red; $bad++ }

Write-Host "[3/3] WPS COM..." -NoNewline
$wps = & $PythonPath -c "import sys, os; sys.path.insert(0, r'$SitePackages\win32'); sys.path.insert(0, r'$SitePackages'); sys.path.insert(0, r'$SitePackages\win32\lib'); os.add_dll_directory(r'$SitePackages\pywin32_system32'); import win32com.client; app = win32com.client.Dispatch('KWPS.Application'); print(app.Name); app.Quit()" 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK" -ForegroundColor Green; $ok++ }
else { Write-Host " WARN (optional)" -ForegroundColor Yellow }

Write-Host ""
if ($bad -eq 0) { Write-Host "All checks passed!" -ForegroundColor Green }
else { Write-Host "$bad failure(s)" -ForegroundColor Red }
