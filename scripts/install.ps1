# wordhelp One-Click Install
# Run: powershell -ExecutionPolicy Bypass -File scripts\install.ps1

param([switch]$Minimal)

$ErrorActionPreference = "Continue"
Write-Host "`n=== wordhelp Installer ===" -ForegroundColor Cyan

# 1. Python + python-docx
Write-Host "[1/4] Python + python-docx..." -NoNewline
$py = & D:\python.exe -c "import docx; print(docx.__version__)" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host " OK ($py)" -ForegroundColor Green
} else {
    Write-Host "`n  Installing python-docx..."
    D:\python.exe -c "import sys; sys.path.insert(0, r'D:\Lib\site-packages'); from pip._internal.cli.main import main; main(['install', 'python-docx'])" 2>&1 | Out-Null
    $py = & D:\python.exe -c "import docx; print(docx.__version__)" 2>&1
    if ($LASTEXITCODE -eq 0) { Write-Host "  OK ($py)" -ForegroundColor Green }
    else { Write-Host "  FAILED - install manually: pip install python-docx" -ForegroundColor Red }
}

# 2. .NET SDK
Write-Host "[2/4] .NET SDK 8.0+..." -NoNewline
$dot = dotnet --version 2>&1
if ($LASTEXITCODE -eq 0) {
    $major = [int]($dot -split '\.')[0]
    if ($major -ge 8) { Write-Host " OK ($dot)" -ForegroundColor Green }
    else { Write-Host " FAILED (need >= 8.0, got $dot)" -ForegroundColor Red }
} else {
    Write-Host "`n  Install .NET SDK 8.0 from: https://dotnet.microsoft.com/download" -ForegroundColor Yellow
    Write-Host "  Or run: winget install Microsoft.DotNet.SDK.8" -ForegroundColor Yellow
}

# 3. WPS COM (optional)
Write-Host "[3/4] WPS Office COM..." -NoNewline
$wps = & D:\python.exe -c "import sys, os; sys.path.insert(0,'D:/Lib/site-packages/win32'); sys.path.insert(0,'D:/Lib/site-packages'); sys.path.insert(0,'D:/Lib/site-packages/win32/lib'); os.add_dll_directory(r'D:\Lib\site-packages\pywin32_system32'); import win32com.client; app=win32com.client.Dispatch('KWPS.Application'); print(app.Name); app.Quit()" 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host " OK" -ForegroundColor Green }
else { Write-Host " WARN (optional - needed for .doc conversion)" -ForegroundColor Yellow }

# 4. minimax-docx (heavy engine)
if (-not $Minimal) {
    Write-Host "[4/4] minimax-docx backend..." -NoNewline
    $skillPath = "$env:USERPROFILE\.codex\skills\minimax-docx"
    if (Test-Path $skillPath) {
        Write-Host " found" -ForegroundColor Green
        Write-Host "  Building backend..."
        $buildDir = "$env:TEMP\minimax-docx-build"
        Remove-Item $buildDir -Recurse -Force -ErrorAction SilentlyContinue
        robocopy "$skillPath\scripts\dotnet" $buildDir /E /NFL /NDL /NJH /NJS /NC | Out-Null
        dotnet restore "$buildDir\MiniMaxAIDocx.Cli\MiniMaxAIDocx.Cli.csproj" --verbosity quiet
        dotnet build "$buildDir\MiniMaxAIDocx.Cli\MiniMaxAIDocx.Cli.csproj" -c Release --no-restore --verbosity quiet
        if ($LASTEXITCODE -eq 0) { Write-Host "  Build OK" -ForegroundColor Green }
        else { Write-Host "  Build FAILED" -ForegroundColor Red }
    } else {
        Write-Host "`n  minimax-docx skill not found." -ForegroundColor Yellow
        Write-Host "  Install it in Codex/Trae skill marketplace: 'Word 文档生成' by MiniMaxAI" -ForegroundColor Yellow
        Write-Host "  Or run with -Minimal to skip (python-docx only mode)" -ForegroundColor Yellow
    }
}

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Run smoke-test to verify: powershell scripts\smoke-test.ps1"
