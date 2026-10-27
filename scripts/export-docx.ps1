# vn-bizdoc: orchestrated DOCX export (Word automation preferred, Flat OPC fallback)
param(
  [Parameter(Mandatory=$true)][string]$Html,
  [string]$Out,
  [ValidateSet('auto','saveas','flat')][string]$Mode = 'auto',
  [int]$SaveAsTimeoutSec = 90
)
$ErrorActionPreference = 'Stop'
$scriptDir = $PSScriptRoot

function Stop-AutomationWord {
  Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match '-Embedding' } |
    ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
}

function Test-FileLocked([string]$Path) {
  if (-not (Test-Path $Path)) { return $false }
  try { $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $fs.Close(); return $false }
  catch { return $true }
}

function Wait-Worker([string]$LogPath, [int]$TimeoutSec) {
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
    Start-Sleep -Seconds 3
    if (Test-Path $LogPath) {
      $c = Get-Content $LogPath -Raw -ErrorAction SilentlyContinue
      if ($c -match 'DONE') { return 'done' }
      if ($c -match 'FATAL') { return 'fatal' }
    }
  }
  return 'timeout'
}

$HtmlPath = (Resolve-Path -LiteralPath $Html).Path
if (-not $Out) { $Out = [System.IO.Path]::ChangeExtension($HtmlPath, '.docx') }
$OutPath = [System.IO.Path]::GetFullPath($Out)

# Fail fast if the target is open in Word (or any other app) - Word cannot overwrite an open file.
if (Test-FileLocked $OutPath) {
  throw ("Output file is open in another application (probably Word). Close it and re-run: " + $OutPath)
}

$work = Join-Path $env:TEMP 'vn-bizdoc'
New-Item -ItemType Directory -Path $work -Force | Out-Null
$logSave = Join-Path $work 'saveas.log'
$logFlat = Join-Path $work 'flat.log'
$flatXml = Join-Path $work 'flat.xml'

Stop-AutomationWord
Start-Sleep -Seconds 1

$attempt1Done = $false
if ($Mode -eq 'saveas' -or $Mode -eq 'auto') {
  Write-Output 'Attempt 1: Word automation SaveAs...'
  Remove-Item $OutPath -Force -ErrorAction SilentlyContinue
  Remove-Item $logSave -Force -ErrorAction SilentlyContinue
  $args1 = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f (Join-Path $scriptDir '_worker-saveas.ps1')),'-Src',('"{0}"' -f $HtmlPath),'-Dst',('"{0}"' -f $OutPath),'-Log',('"{0}"' -f $logSave))
  Start-Process powershell -ArgumentList $args1 -WindowStyle Hidden
  $r1 = Wait-Worker $logSave $SaveAsTimeoutSec
  Stop-AutomationWord
  if ($r1 -eq 'done' -and (Test-Path $OutPath)) {
    Write-Output ('Word SaveAs: OK -> ' + $OutPath)
    $attempt1Done = $true
  } elseif ($Mode -eq 'saveas') {
    throw ("Word SaveAs failed ($r1). See log: " + $logSave)
  } else {
    Write-Output ('Word SaveAs unavailable (' + $r1 + ') - falling back to WordOpenXML + OPC pack')
  }
}

if (-not $attempt1Done -and $Mode -ne 'saveas') {
  Write-Output 'Attempt 2: WordOpenXML serialization + Flat OPC pack...'
  if (Test-FileLocked $OutPath) {
    throw ("Output file is open in another application (probably Word). Close it and re-run: " + $OutPath)
  }
  Stop-AutomationWord
  Remove-Item $logFlat -Force -ErrorAction SilentlyContinue
  Remove-Item $flatXml -Force -ErrorAction SilentlyContinue
  $args2 = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f (Join-Path $scriptDir '_worker-flat.ps1')),'-Src',('"{0}"' -f $HtmlPath),'-Flat',('"{0}"' -f $flatXml),'-Log',('"{0}"' -f $logFlat))
  Start-Process powershell -ArgumentList $args2 -WindowStyle Hidden
  $r2 = Wait-Worker $logFlat 180
  Stop-AutomationWord
  if ($r2 -ne 'done' -or -not (Test-Path $flatXml)) {
    throw ("WordOpenXML fallback failed ($r2). See log: " + $logFlat)
  }
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scriptDir 'pack-flat-opc.ps1') -Flat $flatXml -Out $OutPath
  Write-Output ('Flat OPC pack: OK -> ' + $OutPath)
}

if (-not (Test-Path $OutPath)) { throw 'Export finished without producing the output file.' }

# --- Verification (never fatal: warn if the freshly written file cannot be read) ---
Stop-AutomationWord
try {
  $htmlTxt = [System.IO.File]::ReadAllText($HtmlPath)
  $linksHtml = ([regex]::Matches($htmlTxt, '<a href=')).Count
  $tblHtml = ([regex]::Matches($htmlTxt, '<table')).Count
  Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
  $zip = [System.IO.Compression.ZipFile]::OpenRead($OutPath)
  $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
  $reader = New-Object System.IO.StreamReader($entry.Open())
  $xmlDoc = $reader.ReadToEnd()
  $reader.Close(); $zip.Dispose()
  $linksDocx = ([regex]::Matches($xmlDoc, '<w:hyperlink')).Count
  $tblDocx = ([regex]::Matches($xmlDoc, '<w:tbl>')).Count
  Write-Output ("Parity check -> links: html=$linksHtml docx=$linksDocx | tables: html=$tblHtml docx=$tblDocx")
} catch {
  Write-Output ('Verification skipped: ' + $_.Exception.Message)
}

if (Get-Command officecli -ErrorAction SilentlyContinue) {
  try {
    & officecli close $OutPath 2>&1 | Out-Null
    Start-Sleep -Seconds 1
    & officecli validate $OutPath 2>&1 | Select-Object -First 3
    & officecli close $OutPath 2>&1 | Out-Null
  } catch { Write-Output ('officecli check skipped: ' + $_.Exception.Message) }
} else {
  Write-Output 'officecli not found - skipped validate (install officecli for full verification)'
}
Write-Output ('DONE: ' + $OutPath)
