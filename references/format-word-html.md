# Format: Word-ready HTML (Vietnamese documents)

The single HTML file IS the source of truth. Word/PDF are generated from it. Never edit the .docx.

## 1. Base formatting (match `assets/template.html`)

- `@page { size: A4; margin: 2cm; }`
- Body: Times New Roman, 12pt, line-height 1.45; tables 10.5pt; `.note` 10–10.5pt italic.
- h1 17pt centered; h2 14pt with bottom border; h3 12.5pt; h4 12pt.
- Utility classes: `.center` (numbers/short cells), `.note` (footnotes/disclaimers), `.ref` (source column in price tables), `.docinfo` (info table), `a { color:#1155cc }`.
- List indentation: set indentation with `padding-left` on `ul/ol` (16-18px), never `margin-left`. Word maps a left margin to paragraph indentation on top of the built-in bullet indent, which renders bullets far too deep in the .docx; with margin 0 Word applies its own default (~0.6cm).
- UTF-8 **with BOM**. If the editor tool doesn't add BOM, fix with:
  `$t=[IO.File]::ReadAllText($p); [IO.File]::WriteAllText($p,$t,(New-Object Text.UTF8Encoding($true)))`

## 2. Vietnamese typography

- Decimal comma: `4,5` · thousands dot: `35.990.000đ` · dates `dd/mm/yyyy`.
- Money: `triệu đồng` (or `VNĐ`/`đ`) — be consistent; USD amounts like `99 USD/năm`.
- Effort unit: **"ngày công"** (1 ngày công ≈ 8 giờ; một tháng ≈ 21–22 ngày công). Never use "người-tuần"/"người-tháng" unless the user demands it.
- Heading style: sentence-ish with numbering `1.`, `1.1`, `4.2`; use CAPS only for h1/h2 titles if that matches the existing doc.
- Avoid em-dash overuse; prefer " — " sparingly, colon lists for enumerations.
- Per-project banned words: keep a list (e.g. "tóm tắt" when banned, "bản ngắn gọn" self-labels, "Capacitor" pre-decision) and grep for them after every round.

## 3. Links (clickable in Word)

- EVERY URL in the HTML must be a real `<a href="https://...">` anchor — plain text URLs are invisible as links in Word.
- Trailing punctuation stays OUTSIDE the anchor: `...</a>.` not `....</a>`
- In price-table `.ref` cells, short domains are fine as display text; the href gets `https://`:
  `<a href="https://cellphones.com.vn/...">cellphones.com.vn/...</a> [31]`
- The `.ref` source column should contain clickable links, not only `[n]` refs.

## 4. Document skeleton (default)

1. `<h1>` title (+ optional italic subtitle; no "summary/degraded version" labels).
2. Info table (`docinfo`): Ngày lập, Trạng thái (+ version if multi-round). Do NOT reference other documents here unless explicitly requested.
3. **Lịch sử phiên bản** table: `Phiên bản | Ngày | Nội dung thay đổi` — one row per review round, short and factual.
4. TÓM TẮT/KIẾN NGHỊ or PID structure (Bối cảnh → Mục tiêu → Phạm vi → Phương án → Luận chứng kinh tế → Tổ chức → Kế hoạch → Chất lượng → Rủi ro → Kiểm soát).
5. Appendices: A) prices/sources, B) claim sources (+ glossary if useful).

## 5. Tables

- First column groupings by label; numeric/short cells use `class="center"`.
- Price tables: `Sản phẩm (cấu hình) | Giá niêm yết | Giá bán tham khảo | Nơi bán | Nguồn tham khảo` (last col = `.ref` with links).
- Option tables: `Phương án | Sản phẩm | Giá (tra ngày) | Phù hợp ở điểm nào`.
- Add a closing `.note`: "Giá tra ngày …; cần lấy báo giá chính thức tại thời điểm mua."

## 6. Versioning & review rounds

- Every feedback round: apply targeted edits → bump `Phiên bản` in docinfo AND changelog row (không quên) → sync numbers/decision text across ALL files touched by the same feedback.
- If multiple documents exist (full + presentation): update cross-references (e.g. "phiên bản 1.X" mentions) in the same round.
- Keep changelog rows honest but brief ("Thay đơn vị ước lượng sang ngày công; chốt cấu hình Mac mini M4…").

## 7. Companion documents (full vs presentation)

- A presentation version must be **standalone**: no "bản ngắn gọn", no "xem bản đầy đủ", no reference to another version existing.
- Same numbers/decisions as the full version; shorter prose, keep: summary + recommendations, key rationale tables, estimates, budget, risks, next steps, and a compact source appendix.
- File naming: avoid words that reveal "summary/degraded" status when the user asks for that (e.g. use audience-based names).

## 8. Editing hygiene (run after EVERY round)

- grep banned words & obsolete terms (old names, old units, superseded decisions)
- validate numbers cross-document (totals = sum of parts; version refs consistent)
- check HTML tag balance and UTF-8 BOM
- recount links (`<a href=`) and tables to compare later with exported docx
- keep the docinfo/changelog in sync with the exports you regenerate

## 9. Tone & register (leadership-ready Vietnamese)

- Use formal administrative Vietnamese (văn phong hành chính): impersonal, declarative, decision-oriented. Address the reader as "Ban lãnh đạo"; open proposals with "Kính trình ... xem xét, phê duyệt" and close with "kính đề nghị".
- Preferred verbs: đề nghị, phê duyệt, bảo đảm, triển khai, thực hiện, ban hành. Avoid colloquialisms ("cần mua luôn", "khá là", "nhé") and machine-flavoured filler ("nhằm hướng tới việc", "một cách toàn diện", "đóng vai trò quan trọng trong việc").
- No invented compounds, no word-by-word translation of English jargon. Keep the official term once with a Vietnamese gloss at first use - "gói phân phối đầy đủ (Full distribution)", "người đại diện tài khoản (Account Holder)" - then use the Vietnamese term. If no settled Vietnamese term exists, keep the English term in parentheses after a plain description.
- Write bullets as full sentences; avoid templated "**Nhãn:** chi tiết" lists outside tables.
- No em dashes and no tilde for approximations: write "khoảng 3,26 triệu đồng", not "~3,26 triệu".
- When the user asks to "humanize" or flags the tone, run the `humanizer` skill's rules on the draft (no em dashes, no bold-lead lists, no summary-style conclusions, no invented terms).
