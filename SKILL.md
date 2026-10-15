---
name: vn-bizdoc
description: End-to-end workflow for Vietnamese business documents (đề xuất, báo cáo, proposals/reports for leadership): sourcing research with numbered citations, Word-ready HTML formatting (A4, Times New Roman, clickable links, appendices, version history), QA checks, and export to .docx (Word automation preferred, with a verified fallback) or .pdf (headless browser print). Use when the user asks to viết báo cáo / viết đề xuất / tài liệu gửi lãnh đạo, prepare a proposal in Vietnamese, export to docx/Word or PDF, or convert a styled Vietnamese HTML document to Word.
---

# vn-bizdoc — Vietnamese Business Documents: Research → HTML → DOCX/PDF

Process proven on real leadership-facing proposals (pricing research, Word export trouble-shooting, multi-round review). Documents are written in **Vietnamese**; this skill (instructions) is in English on purpose.

## When to use

- User wants a professional Vietnamese document: đề xuất, báo cáo, tờ trình, tài liệu trình lãnh đạo, ước lượng/ngân sách, phân tích mua sắm.
- Output must look right in Word: A4, Times New Roman, tables, clickable links, source appendices.
- User asks to "xuất docx / chuyển sang Word / xuất PDF".

## Golden rules (do not skip)

1. **Primary deliverable = ONE self-contained styled HTML file** (UTF-8 with BOM). DOCX/PDF are derivatives.
2. **Ask before exporting.** Unless the user explicitly asked for export in the current request, finish and present the HTML first, then ask: *"Xuất luôn .docx/.pdf hay để rà soát xong mới xuất?"* Default format when asked to export without specifics: `.docx`; offer `.pdf`.
3. **Every external figure needs a source + date:** prices, fees, policies, specs, legal references. Numbered `[n]` citations in the body; full lists in appendices (A: prices/devices/accounts; B: sources for key claims). Close price sections with: *"Giá tra ngày dd/mm/yyyy; cần kiểm tra lại khi mua."*
4. **Each review round = version bump + changelog row** in the document's "Lịch sử phiên bản" table, and sync companion files (full version vs presentation version).
5. **Never leak document relationships** the user didn't ask for: a presentation doc must not say "bản ngắn gọn" or reference "bản đầy đủ"; each document stands alone.
6. Respect the project's vocabulary decisions — e.g. effort unit is **"ngày công"** (never "người-tuần"/"người-tháng" unless asked); keep a per-project banned-words list and grep for it after every edit round.
7. HTML is the source of truth. Edit HTML, then re-export; never edit the .docx directly.

## Workflow

### Phase 0 — Clarify (keep it short)
Confirm: audience, purpose, length (e.g. full vs 3–5 page presentation version), structure (PID/PRINCE2-style, proposal, report), budget/technical scope, and any banned/required terms. Do not over-ask; sensible defaults below.

### Phase A — Research & sourcing
- Extract the "facts requiring sources" list from the brief (prices, fees, timeline facts, platform policies, legal refs, "why build vs buy" claims).
- Research **in parallel**; prefer official sellers/pages (see `references/research-sourcing.md` for the Vietnam source shortlist and capture rules).
- Normalize each fact to `{value, unit, source, date}` and assign `[n]` numbers used everywhere.
- Build Appendix A (prices) and Appendix B (claim sources) from the start.

### Phase B — Write & format (HTML)
- Start from `assets/template.html` (keeps the exact CSS: A4 2cm, Times New Roman 12pt, table 10.5pt, `.note`, `.ref`, `.center`).
- Follow `references/format-word-html.md` for: Vietnamese typography, link rules (all URLs are `<a href>`), table/source-column conventions, version history + changelog protocol, companion-document policy, banned-words QA.
- Run the QA checklist (`assets/qa-checklist.md`) after every editing round.

### Phase C — Export (.docx / .pdf) — only on user request
Preferred: **Word automation** via `scripts/export-docx.ps1` (it auto-falls back when the machine's Word SaveAs is broken — see `references/word-automation-notes.md`).

```powershell
# DOCX (Word automation preferred; auto fallback WordOpenXML->OPC pack)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\export-docx.ps1 -Html "docs\De-xuat-X.html"
# PDF (headless Edge/Chrome print; Word ExportAsFixedFormat fallback)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\export-pdf.ps1 -Html "docs\De-xuat-X.html"
```

- `-Out` overrides the output path (default: same folder, same name, new extension).
- `-Mode saveas|flat` / `edge|word` force or skip a path (debugging).
- Always finish with verification (script prints it): `officecli validate` = pass; hyperlink/table counts in docx == counts in HTML; Vietnamese text readable at head/tail; no leftover `WINWORD /Automation -Embedding` processes.
- If the user converted manually in Word and the HTML changed afterwards, remind them to re-export.

## Files

- `references/research-sourcing.md` — sourcing method, VN source shortlist, citation/appendices patterns.
- `references/format-word-html.md` — HTML/CSS rules, Vietnamese typography, versioning & multi-file sync.
- `references/word-automation-notes.md` — Word COM pitfalls (zombie processes, SaveAs deadlock), fallback pipeline, PDF options, verification commands.
- `scripts/export-docx.ps1` — orchestrator (Word SaveAs → fallback Flat OPC pack → verification).
- `scripts/_worker-saveas.ps1`, `scripts/_worker-flat.ps1` — background workers it spawns.
- `scripts/pack-flat-opc.ps1` — Word Flat-OPC XML → .docx (synthesizes `[Content_Types].xml`).
- `scripts/export-pdf.ps1` — headless browser print → PDF, Word fallback.
- `assets/template.html` — ready-to-copy HTML skeleton + CSS.
- `assets/qa-checklist.md` — pre-delivery verification checklist.
