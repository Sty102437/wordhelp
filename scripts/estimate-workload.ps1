param([int]$P, [int]$T, [int]$C, [int]$I, [string]$Type="general", [switch]$TOC)
$s = 0
$s += [Math]::Min($P/10, 5)
$s += [Math]::Min($C/5, 5)
$s += [Math]::Min($I/3, 5)
if ($TOC) { $s += 2 }
if ($Type -in @("academic","government")) { $s += 3 }
if ($s -gt 8) { "engine=minimax-docx; score=$s" }
else { "engine=python-docx; score=$s" }
