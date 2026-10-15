# Word Automation & Export Notes (hard-won lessons)

Export scripts: `scripts/export-docx.ps1` (orchestrator), `scripts/_worker-saveas.ps1`, `scripts/_worker-flat.ps1`, `scripts/pack-flat-opc.ps1`, `scripts/export-pdf.ps1`.

## 1. DOCX pipeline (preferred: Word automation)

Order of attempts:

1. **Word COM → `SaveAs2(docx, 16)`** (the normal path; works on most machines).
   - `Visible=$false`, `DisplayAlerts=0`
   - `Documents.Open(html, ConfirmConversions=$false, ReadOnly=$false, AddToRecentFiles=$false)` — opening ReadOnly can trigger save-copy dialogs; avoid.
2. **Fallback when SaveAs hangs** (observed on some machines): `Document.WordOpenXML` (in-memory Flat OPC serialization — no Save API involved) → `pack-flat-opc.ps1` builds the .docx zip itself.

The orchestrator automates this: it runs attempt 1 as a background worker with a log file and polls; on timeout it kills Word and runs the fallback. NEVER call SaveAs2 synchronously in a foreground script — if it deadlocks you can never abort it.

## 2. Environment pitfalls

- **Zombie processes**: crashed/timed-out automation leaves `WINWORD.EXE /Automation -Embedding` running (no window), which keeps file locks. Always sweep before and after:
  ```powershell
  Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
    Where-Object { $_.CommandLine -match '-Embedding' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
  ```
  Only kill `-Embedding` (automation) ones; never kill a user's visible Word.
- **After force-killing Word**, the next start can hang invisibly on the Document Recovery prompt. Fix: export your settings first, then
  ```powershell
  Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency" -Recurse -Force
  ```
  (Word recreates it automatically.)
- **SaveAs2 deadlock symptom**: hangs forever with NO dialog window (verify with a user32 EnumWindows dump), independent of destination folder (even a trivial new doc saved to %TEMP% hangs). This is a machine/Word-profile problem — do not keep retrying; fall back to WordOpenXML.
- **officecli resident** keeps files open (causing "file is being used by another process" in other tools). Run `officecli close <file>` before external readers.
- Word AddIns/COMAddIns are usually NOT the cause (verified) — don't disable them.
- Cleaning a hang: `Stop-Process` the automation PIDs, sweep, then proceed. All Word work must be wrapped in try/finally that quits Word.

## 3. Flat OPC pack (fallback) — key detail

Word's `WordOpenXML` output is a Flat OPC XML where the **`[Content_Types].xml` part is NOT included** — each part carries `pkg:contentType` attributes instead. The packer must synthesize `[Content_Types].xml`:

- Defaults: `rels` → relationship content type, `xml` → application/xml
- One `<Override PartName="/word/..." ContentType="<pkg:contentType>"/>` per non-`.rels` part
- Write `[Content_Types].xml` as the FIRST zip entry, then all parts (`xmlData` child serialized via XmlWriter; `binaryData` base64-decoded).
- Use `Documents.Add()` + copy `Content.FormattedText` + copy Section PageSetup (A4) from the HTML-opened doc — then serialize THAT doc (a native Word doc) via `WordOpenXML`. Hyperlinks and tables survive the FormattedText copy (verified: counts match HTML 1:1).

## 4. Verification (always run after export)

```powershell
officecli validate "<out.docx>"          # expect: Validation passed
# hyperlink/table parity vs HTML:
#   count '<a href=' and '<table' in HTML  ==  '<w:hyperlink' and '<w:tbl>' in docx word/document.xml
# OPC sanity (.NET): [System.IO.Packaging.Package]::Open(...)   (after: officecli close)
# text endpoints: officecli view <out.docx> text --max-lines N  (start shows title; tail shows appendix)
# no leftover:  (Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'").Count
```

## 5. PDF export

- Primary: **headless Edge/Chrome** `--headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf=<out> <file-url>` — reproduces the HTML design exactly (the `@page` CSS already sets A4/2cm margins) and keeps links clickable. Retry with plain `--headless` if the flag variant fails.
- Fallback: Word COM `ExportAsFixedFormat(<out>, 17)` (wdExportFormatPDF) — may hang on machines with the SaveAs deadlock; only use when no browser found.
