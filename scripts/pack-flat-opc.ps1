# vn-bizdoc: pack a Word Flat OPC XML (from Document.WordOpenXML) into a valid .docx
# Key detail: Word's Flat OPC omits the [Content_Types].xml part - synthesize it from each part's pkg:contentType attribute.
param(
  [Parameter(Mandatory=$true)][string]$Flat,
  [Parameter(Mandatory=$true)][string]$Out
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

$pns = 'http://schemas.microsoft.com/office/2006/xmlPackage'
$xdoc = New-Object System.Xml.XmlDocument
$xdoc.LoadXml([System.IO.File]::ReadAllText($Flat, [System.Text.Encoding]::UTF8))
$nsm = New-Object System.Xml.XmlNamespaceManager($xdoc.NameTable)
$nsm.AddNamespace('pkg', $pns)
$parts = @($xdoc.SelectNodes('//pkg:part', $nsm))
if ($parts.Count -eq 0) { throw 'No pkg:part elements found - not a Flat OPC file?' }

# Synthesize [Content_Types].xml
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$sb.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
[void]$sb.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
[void]$sb.Append('<Default Extension="xml" ContentType="application/xml"/>')
$hasCT = $false
foreach ($part in $parts) {
  $n = $part.GetAttribute('name', $pns)
  if ($n -eq '/[Content_Types].xml') { $hasCT = $true; continue }
  if ($n -match '\.rels$') { continue }
  $ctv = $part.GetAttribute('contentType', $pns)
  if ($ctv) { [void]$sb.Append('<Override PartName="' + $n + '" ContentType="' + $ctv + '"/>') }
}
[void]$sb.Append('</Types>')

$OutPath = [System.IO.Path]::GetFullPath($Out)
Remove-Item $OutPath -Force -ErrorAction SilentlyContinue
$fs = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::CreateNew)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  if (-not $hasCT) {
    $entry = $zip.CreateEntry('[Content_Types].xml', [System.IO.Compression.CompressionLevel]::Optimal)
    $es = $entry.Open()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $es.Write($bytes, 0, $bytes.Length)
    $es.Close()
  }
  foreach ($part in $parts) {
    $name = $part.GetAttribute('name', $pns).TrimStart('/')
    $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $es = $entry.Open()
    try {
      $xmlData = $part.SelectSingleNode('pkg:xmlData', $nsm)
      $binData = $part.SelectSingleNode('pkg:binaryData', $nsm)
      if ($xmlData -ne $null -and $xmlData.FirstChild -ne $null) {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $false
        $w = [System.Xml.XmlWriter]::Create($es, $settings)
        try { $xmlData.FirstChild.WriteTo($w); $w.Flush() } finally { $w.Close() }
      } elseif ($binData -ne $null) {
        $bytes2 = [Convert]::FromBase64String($binData.InnerText)
        $es.Write($bytes2, 0, $bytes2.Length)
      }
    } finally { $es.Close() }
  }
} finally { $zip.Dispose(); $fs.Close() }
Write-Output ('Packed: ' + $OutPath + ' (' + (Get-Item $OutPath).Length + ' bytes, ' + $parts.Count + ' parts)')
