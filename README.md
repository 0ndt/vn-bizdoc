# vn-bizdoc

An **agent skill** for producing professional **Vietnamese business documents** end-to-end — proposals, reports, and submissions to leadership:

**source-backed research → Word-ready HTML → export to DOCX / PDF (on demand).**

Built and battle-tested on real leadership-facing proposals (pricing research across Vietnamese retailers, multi-round reviews, Word export troubleshooting).

[Vietnamese version](README.vi.md)

---

## What it does

### Phase A — Research & sourcing
- Extracts the "facts requiring sources" list from the brief (prices, fees, platform policies, legal references, build-vs-buy claims).
- Prefers official sources (Apple, LG, Google, GitHub) and major Vietnamese retailers (CellphoneS, GearVN, FPT Shop, TGDD, TopZone, Hacom…), with capture rules for list price vs sale price, discontinued products, and date stamping.
- Numbered `[n]` citations used throughout; source lists live in appendices: **A — prices/devices/accounts**, **B — sources for key claims**.

### Phase B — Write & format (HTML is the master file)
- Word-ready single-file HTML: A4 · 2 cm margins · Times New Roman 12pt · styled tables · `.note` / `.ref` / `.center` utilities.
- Vietnamese typography rules (decimal comma, dd/mm/yyyy dates, man-day effort unit, per-project banned-words list).
- **Every URL is a clickable `<a href>`** — survives Word import as real hyperlinks.
- Document skeleton with **version history + changelog**; multi-file sync (full version vs standalone presentation version); editing-round protocol.

### Phase C — Export (only when the user asks)
- **DOCX** — Word automation preferred (`SaveAs2`), with an automatic, verified fallback for machines where Word's SaveAs deadlocks: `WordOpenXML` serialization + **Flat OPC pack** (synthesizes `[Content_Types].xml`).
- **PDF** — headless Edge/Chrome print (keeps the HTML design and clickable links); Word `ExportAsFixedFormat` fallback.
- Built-in verification: `officecli validate`, hyperlink/table parity vs the HTML, zombie Word process cleanup.

---

## Repository layout

```
vn-bizdoc/
├── SKILL.md                        # workflow + golden rules (agent entry point)
├── README.md                       # this file (English)
├── README.vi.md                    # Vietnamese README
├── LICENSE                         # MIT
├── references/
│   ├── research-sourcing.md        # Phase A: sourcing method, source shortlist, citation patterns
│   ├── format-word-html.md         # Phase B: CSS/typography rules, versioning, multi-file sync
│   └── word-automation-notes.md    # Phase C: Word COM pitfalls, fallback pipeline, verification
├── scripts/
│   ├── export-docx.ps1             # orchestrator: SaveAs → fallback → verification
│   ├── _worker-saveas.ps1          # background worker (timeout-detectable Word SaveAs)
│   ├── _worker-flat.ps1            # background worker (WordOpenXML, no Save API)
│   ├── pack-flat-opc.ps1           # Flat OPC XML → valid .docx
│   └── export-pdf.ps1              # headless browser print → PDF
└── assets/
    ├── template.html               # ready-to-copy HTML skeleton + CSS
    └── qa-checklist.md             # pre-delivery QA checklist
```

---

## Installation

Copy this folder into a skills directory recognized by your agent framework:

| Scope | Path (Windows) | Path (macOS/Linux) |
|---|---|---|
| User-level (all projects) | `C:\Users\<you>\.agents\skills\vn-bizdoc` | `~/.agents/skills/vn-bizdoc` |
| Project-level | `<repo>\.agents\skills\vn-bizdoc` | `<repo>/.agents/skills/vn-bizdoc` |

Restart the agent session so the skill gets discovered.

## Usage

Ask the agent to write a document, for example:

- "Use vn-bizdoc to write a device-purchase proposal for the development team."
- "Use vn-bizdoc to prepare a quarterly report for leadership."

The skill researches sources, writes the styled HTML, and then asks: *"Export to .docx/.pdf now, or after review?"*

Export manually at any time:

```powershell
# DOCX (Word automation preferred; auto fallback when SaveAs hangs)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\export-docx.ps1 -Html "docs\proposal.html"

# PDF (headless Edge/Chrome; Word fallback)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\export-pdf.ps1 -Html "docs\proposal.html"
```

Options: `-Out <path>` (default: same folder, new extension) · `-Mode saveas|flat` (DOCX path) · `-Mode edge|word` (PDF path) · `-SaveAsTimeoutSec N`.

Each export ends with verification output:

```
Parity check -> links: html=16 docx=16 | tables: html=6 docx=6
Validation passed: no errors found.
DONE: docs\proposal.docx
```

## Requirements

| Requirement | Needed for |
|---|---|
| Windows + Microsoft Word | DOCX export (both the primary and fallback paths) |
| PowerShell 5.1+ | all scripts |
| Microsoft Edge or Google Chrome | PDF export (primary path) — Word fallback otherwise |
| [officecli](https://d.officecli.ai) (optional) | automatic OpenXML validation of exported files |

## How the DOCX fallback works

Some machines hit a Word automation deadlock where `SaveAs2` hangs indefinitely with no dialog (verified: destination-independent, even for a trivial save). The orchestrator detects this via a timeout on a background worker, kills Word, then:

1. Opens the HTML in Word and copies content into a native document (hyperlinks and tables survive — verified 1:1).
2. Serializes the document in memory via `Document.WordOpenXML` (**no Save API involved**).
3. Packs the Flat OPC XML into a valid `.docx` with `pack-flat-opc.ps1` (Word's Flat OPC omits `[Content_Types].xml`; the script synthesizes it from each part's `pkg:contentType`).

See [`references/word-automation-notes.md`](references/word-automation-notes.md) for the full troubleshooting notes.

## License

MIT — see [LICENSE](LICENSE).
