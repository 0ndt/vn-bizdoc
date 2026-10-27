# vn-bizdoc QA Checklist (run before delivering / exporting)

## Content & language
- [ ] Every external figure has `[n]` citation + source in Phụ lục A/B; price closing line present ("Giá tra ngày …; cần kiểm tra lại khi mua.")
- [ ] Citation placement audit: each `[n]` sits immediately after the exact sourced fragment, not at the end of a sentence that also carries internal estimates or analysis; vendor-specific claims cite only that vendor's source.
- [ ] Citation numbering is contiguous and maps 1:1 to the reference list.
- [ ] Totals = sum of parts (budgets, estimates, ngày công) — recompute after every edit
- [ ] Version bumped + changelog row added for this round
- [ ] Companion docs synced (same numbers/decisions; version refs updated)
- [ ] Standalone rule respected (no "bản ngắn gọn"/"bản đầy đủ"/cross-doc references unless asked)
- [ ] Banned/obsolete words grep = 0 (project list): e.g. `người-tuần`, `người-tháng`, `TÓM TẮT` (if banned), old brand names, superseded decisions
- [ ] Vietnamese typography: decimal commas, dd/mm/yyyy, consistent đ/đồng

## Structure & HTML
- [ ] Tag balance check (table/tr/td/th/ul/ol/li/p/h2/h3 + `a` open/close counts equal)
- [ ] Every URL is `<a href>`; trailing punctuation outside anchors; no double-wrapped anchors
- [ ] UTF-8 BOM present:
  `$b=[IO.File]::ReadAllBytes($p)[0..2]; "{0:X2} {1:X2} {2:X2}" -f $b[0],$b[1],$b[2]`  → must start `EF BB BF`
  (fix: `$t=[IO.File]::ReadAllText($p); [IO.File]::WriteAllText($p,$t,(New-Object Text.UTF8Encoding($true)))`)
- [ ] Links/table counts recorded (for export parity): `([regex]::Matches($c,'<a href=')).Count`, `([regex]::Matches($c,'<table')).Count`

## Export (only when user asked)
- [ ] Ran `scripts/export-docx.ps1` (or `export-pdf.ps1`); output path reported to user
- [ ] Parity printed by script: html links == docx `<w:hyperlink`; html tables == docx `<w:tbl>`
- [ ] `officecli validate` = "Validation passed"; then `officecli close` the file
- [ ] Head/tail text readable Vietnamese (`officecli view <f> text --max-lines N` / tail) — title at top, appendix at bottom
- [ ] No leftover automation Word: `@(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'").Count` → 0 (or only user's visible instance)
- [ ] Stale exports removed/flagged (old docx from earlier review rounds)

## Housekeeping
- [ ] Temp junk removed (`%TEMP%\vn-bizdoc`, failed artifacts like `<name>.html.docx`)
- [ ] User informed what changed + what remains optional (e.g. old files that can be deleted)
- [ ] No commit unless user asked
