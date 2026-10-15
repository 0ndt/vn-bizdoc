# vn-bizdoc worker: open HTML in Word and SaveAs2 as .docx (runs in a background process; parent polls the log)
param(
  [Parameter(Mandatory=$true)][string]$Src,
  [Parameter(Mandatory=$true)][string]$Dst,
  [Parameter(Mandatory=$true)][string]$Log
)
Set-Content -Path $Log -Value ('START ' + (Get-Date))
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  Add-Content $Log 'word-created'
  try {
    $doc = $word.Documents.Open($Src, $false, $false, $false)
    Add-Content $Log 'opened'
    $doc.SaveAs2($Dst, 16)   # 16 = wdFormatDocumentDefault (.docx)
    Add-Content $Log ('saved size=' + (Get-Item $Dst).Length)
    $doc.Close(0)
  } finally {
    try { $word.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
  }
  Add-Content $Log 'DONE'
} catch {
  Add-Content $Log ('FATAL: ' + $_.Exception.Message)
}
