# Open the given docx files in real Word and read their statistics.
# ASCII-only on purpose: Windows PowerShell 5.1 reads .ps1 as ANSI, so any
# non-ASCII here would be mangled and break the parser.
#
# This is the one step nothing else can prove: a zip that unpacks is not
# the same as a file Word agrees to open.
#
# Usage: powershell -ExecutionPolicy Bypass -File test\word.ps1 <a.docx> [b.docx ...]

param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Paths)

if (-not $Paths -or $Paths.Count -eq 0) {
  Write-Host 'Usage: powershell -File test\word.ps1 <a.docx> [b.docx ...]'
  exit 1
}

$ErrorActionPreference = 'Stop'
$outFile = Join-Path $PSScriptRoot '_word_check.txt'
$alreadyOpen = @(Get-Process WINWORD -ErrorAction SilentlyContinue).Count
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$lines = @()
try {
  foreach ($p in $Paths) {
    $label = [System.IO.Path]::GetFileName($p)
    if (-not (Test-Path $p)) { $lines += "${label}: MISSING"; continue }
    try {
      $doc = $word.Documents.Open($p, [ref]$false, [ref]$true, [ref]$false)
      $pages  = $doc.ComputeStatistics(2)   # wdStatisticPages
      $words  = $doc.ComputeStatistics(0)   # wdStatisticWords
      $chars  = $doc.ComputeStatistics(3)   # wdStatisticCharacters
      $shapes = $doc.InlineShapes.Count
      $paras  = $doc.Paragraphs.Count
      $text   = $doc.Content.Text
      if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
      $text = $text -replace "`r", ' | ' -replace "`n", ' '
      $lines += "-- $label"
      $lines += "   pages=$pages words=$words chars=$chars paras=$paras images=$shapes"
      $lines += "   text: $text"
      $doc.Close([ref]$false)
    } catch {
      $lines += "-- $label  FAILED: $($_.Exception.Message)"
    }
  }
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  # only clean up stray Word processes if the user had none open to begin with
  if ($alreadyOpen -eq 0) {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  }
}

$lines | Out-File -FilePath $outFile -Encoding utf8
$lines | ForEach-Object { $_ }
