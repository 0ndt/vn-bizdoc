# vn-bizdoc worker: open HTML in Word, copy into a native doc, serialize WordOpenXML (Flat OPC) - NO Save API involved.
param(
  [Parameter(Mandatory=$true)][string]$Src,
  [Parameter(Mandatory=$true)][string]$Flat,
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
    $new = $word.Documents.Add()
    $new.Content.FormattedText = $doc.Content.FormattedText
    try {
      $ss = $doc.Sections.Item(1).PageSetup
      $ts = $new.Sections.Item(1).PageSetup
      $ts.PageWidth = $ss.PageWidth
      $ts.PageHeight = $ss.PageHeight
      $ts.TopMargin = $ss.TopMargin
      $ts.BottomMargin = $ss.BottomMargin
      $ts.LeftMargin = $ss.LeftMargin
      $ts.RightMargin = $ss.RightMargin
    } catch {}
    Add-Content $Log 'serializing'
    $flatXml = $new.WordOpenXML
    [System.IO.File]::WriteAllText($Flat, $flatXml, (New-Object System.Text.UTF8Encoding($false)))
    Add-Content $Log ('flat-saved size=' + (Get-Item $Flat).Length)
    $new.Close(0)
    $doc.Close(0)
  } finally {
    try { $word.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
  }
  Add-Content $Log 'DONE'
} catch {
  Add-Content $Log ('FATAL: ' + $_.Exception.Message)
}
