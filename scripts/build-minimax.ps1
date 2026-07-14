param(
    [string]$SkillPath = $(if ($env:WORDHELP_MINIMAX_SKILL) {
        $env:WORDHELP_MINIMAX_SKILL
    } else {
        "$env:USERPROFILE\.codex\skills\minimax-docx\scripts\dotnet"
    })
)
$buildDir = "$env:TEMP\minimax-docx-build"
Write-Host "Building minimax-docx..." -ForegroundColor Cyan
Write-Host "  Source: $SkillPath"
Remove-Item $buildDir -Recurse -Force -ErrorAction SilentlyContinue
robocopy $SkillPath $buildDir /E /NFL /NDL /NJH /NJS /NC | Out-Null
dotnet restore "$buildDir\MiniMaxAIDocx.Cli\MiniMaxAIDocx.Cli.csproj" --verbosity quiet
dotnet build "$buildDir\MiniMaxAIDocx.Cli\MiniMaxAIDocx.Cli.csproj" -c Release --no-restore --verbosity quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "Done. Backend: $buildDir" -ForegroundColor Green
} else {
    Write-Host "Build FAILED. Check .NET SDK and minimax-docx skill path." -ForegroundColor Red
    Write-Host "  Set custom path via: -SkillPath <path> or env WORDHELP_MINIMAX_SKILL" -ForegroundColor Yellow
}
