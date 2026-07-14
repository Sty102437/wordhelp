$SkillPath = "$env:USERPROFILE\.codex\skills\minimax-docx\scripts\dotnet"
$buildDir = "$env:TEMP\minimax-docx-build"
Write-Host "Building minimax-docx..." -ForegroundColor Cyan
Remove-Item $buildDir -Recurse -Force -ErrorAction SilentlyContinue
robocopy $SkillPath $buildDir /E /NFL /NDL /NJH /NJS /NC | Out-Null
dotnet restore "$buildDir\MiniMaxAIDocx.Cli\MiniMaxAIDocx.Cli.csproj" --verbosity quiet
dotnet build "$buildDir\MiniMaxAIDocx.Cli\MiniMaxAIDocx.Cli.csproj" -c Release --no-restore --verbosity quiet
Write-Host "Done. Backend: $buildDir" -ForegroundColor Green
