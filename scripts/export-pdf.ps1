# vn-bizdoc: export HTML -> PDF. Primary: headless Edge/Chrome print (keeps the HTML design + clickable links). Fallback: Word ExportAsFixedFormat.
param(
  [Parameter(Mandatory=$true)][string]$Html,
  [string]$Out,
  [ValidateSet('auto','edge','word')][string]$Mode = 'auto'
)
$ErrorActionPreference = 'Stop'

function Test-FileLocked([string]$Path) {
  if (-not (Test-Path $Path)) { return $false }
  try { $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $fs.Close(); return $false }
  catch { return $true }
}

$HtmlPath = (Resolve-Path -LiteralPath $Html).Path
if (-not $Out) { $Out = [System.IO.Path]::ChangeExtension($HtmlPath, '.pdf') }
$OutPath = [System.IO.Path]::GetFullPath($Out)

if (Test-FileLocked $OutPath) {
  throw ("Output file is open in another application (PDF viewer?). Close it and re-run: " + $OutPath)
}
Remove-Item $OutPath -Force -ErrorAction SilentlyContinue
$start = (Get-Date).AddSeconds(-2)
$url = ([uri]$HtmlPath).AbsoluteUri

$browser = @(
  (Join-Path ${env:ProgramFiles} 'Microsoft\Edge\Application\msedge.exe'),
  (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
  (Join-Path ${env:ProgramFiles} 'Google\Chrome\Application\chrome.exe'),
  (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe')
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if (($Mode -eq 'auto' -or $Mode -eq 'edge') -and $browser) {
  Write-Output ('Headless print via: ' + $browser)
  foreach ($headlessFlag in @('--headless=new','--headless')) {
    $argsList = @($headlessFlag,'--disable-gpu','--no-pdf-header-footer',('--print-to-pdf="' + $OutPath + '"'),('"' + $url + '"'))
    Start-Process -FilePath $browser -ArgumentList $argsList -Wait -NoNewWindow
    if ((Test-Path $OutPath) -and ((Get-Item $OutPath).LastWriteTime -ge $start)) { break }
  }
  if ((Test-Path $OutPath) -and ((Get-Item $OutPath).LastWriteTime -ge $start)) {
    Write-Output ('OK -> ' + $OutPath + ' (' + (Get-Item $OutPath).Length + ' bytes)')
    return
  }
  if ($Mode -eq 'edge') { throw 'Headless browser produced no fresh PDF' }
  Write-Output 'Headless browser failed - falling back to Word export'
}

# Word fallback (may hang on machines with the SaveAs deadlock; see references/word-automation-notes.md)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($HtmlPath, $false, $false, $false)
  $doc.ExportAsFixedFormat($OutPath, 17)   # 17 = wdExportFormatPDF
  $doc.Close(0)
} finally {
  try { $word.Quit() } catch {}
  try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
  Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match '-Embedding' } |
    ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
}
Write-Output ('OK (Word export) -> ' + $OutPath + ' (' + (Get-Item $OutPath).Length + ' bytes)')
