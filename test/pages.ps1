# Report page count and per-page paragraph split for the given docx files.
# ASCII-only on purpose: Windows PowerShell 5.1 reads .ps1 as ANSI.
# Usage: powershell -ExecutionPolicy Bypass -File test\pages.ps1 <a.docx> [b.docx ...]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Paths)

if (-not $Paths -or $Paths.Count -eq 0) {
  Write-Host 'Usage: powershell -File test\pages.ps1 <a.docx> [b.docx ...]'
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
    try {
      $doc = $word.Documents.Open($p, [ref]$false, [ref]$true, [ref]$false)
      $pages = $doc.ComputeStatistics(2)
      $lines += "== $([System.IO.Path]::GetFileName($p))"
      $lines += "   pages=$pages paragraphs=$($doc.Paragraphs.Count)"

      # Which paragraph starts each page: locate the page break positions
      for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $r = $doc.Paragraphs.Item($i).Range
        $pg = $r.Information(3)   # wdActiveEndPageNumber
        $lines += ("   p{0}  para{1}" -f $pg, $i)
      }
      $doc.Close([ref]$false)
    } catch {
      $lines += "FAILED: $($_.Exception.Message)"
    }
  }
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  if ($alreadyOpen -eq 0) {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  }
}

$lines | Out-File -FilePath (Join-Path $PSScriptRoot '_pages.txt') -Encoding utf8
$lines | ForEach-Object { $_ }
