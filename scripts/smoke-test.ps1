Write-Host "=== wordhelp Smoke Test ===" -ForegroundColor Cyan
$ok = 0; $bad = 0

Write-Host "[1/3] Python + python-docx..." -NoNewline
$py = & D:\python.exe -c 'import docx; print(docx.__version__)' 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK ($py)" -ForegroundColor Green; $ok++ }
else { Write-Host " FAILED" -ForegroundColor Red; $bad++ }

Write-Host "[2/3] .NET SDK..." -NoNewline
$dot = dotnet --version 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK ($dot)" -ForegroundColor Green; $ok++ }
else { Write-Host " FAILED" -ForegroundColor Red; $bad++ }

Write-Host "[3/3] WPS COM..." -NoNewline
$wps = & D:\python.exe -c 'import win32com.client; app = win32com.client.Dispatch("KWPS.Application"); print(app.Name); app.Quit()' 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK" -ForegroundColor Green; $ok++ }
else { Write-Host " WARN (optional)" -ForegroundColor Yellow }

Write-Host ""
if ($bad -eq 0) { Write-Host "All checks passed!" -ForegroundColor Green }
else { Write-Host "$bad failure(s)" -ForegroundColor Red }
