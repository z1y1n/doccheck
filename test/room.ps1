# Measure how much vertical room is left at the bottom of the page.
# Usage: powershell -ExecutionPolicy Bypass -File test\room.ps1 <a.docx> [b.docx ...]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Paths)
if (-not $Paths -or $Paths.Count -eq 0) {
  Write-Host 'Usage: powershell -File test\room.ps1 <a.docx> [b.docx ...]'
  exit 1
}
$ErrorActionPreference = 'Stop'
$alreadyOpen = @(Get-Process WINWORD -ErrorAction SilentlyContinue).Count
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$lines = @()
try {
  foreach ($p in $Paths) {
    if (-not (Test-Path $p)) { $lines += "MISSING: $p"; continue }
    $doc = $word.Documents.Open($p, [ref]$false, [ref]$true, [ref]$false)
    $ps = $doc.PageSetup
    $h = $ps.PageHeight
    $btm = $ps.BottomMargin
    $limit = $h - $btm
    $last = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $y = $last.Information(6)   # wdVerticalPositionRelativeToPage
    $lines += ("{0}" -f [System.IO.Path]::GetFileName($p))
    $lines += ("   pageHeight={0:N1} bottomMargin={1:N1} textLimit={2:N1}" -f $h, $btm, $limit)
    $lines += ("   lastParaTop={0:N1}  roomLeft={1:N1} pt (= {2:N1} lines of 12pt)" -f $y, ($limit - $y), (($limit - $y) / 12))
    $lines += ("   pages={0}" -f $doc.ComputeStatistics(2))
    $doc.Close([ref]$false)
  }
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  if ($alreadyOpen -eq 0) { Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }
}
$lines | Out-File -FilePath (Join-Path $PSScriptRoot '_room.txt') -Encoding utf8
$lines | ForEach-Object { $_ }
