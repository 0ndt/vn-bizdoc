# Research & Sourcing

Goal: every external number in the document is traceable — `{value, source URL, date checked}` — and the document carries the citations itself.

## 1. Build the "facts requiring sources" list

From the brief, enumerate categories, e.g.:

- Device/hardware prices (specific configs: "Mac mini M4 24GB/512GB" — include RAM/SSD).
- Account/service fees (Apple Developer, Google Play, Firebase, CI services).
- Platform policies (review times, SDK requirements, store guidelines).
- Legal references (laws, decrees — VN).
- "Why build vs buy" claims (archived projects, official docs, guidelines).

Internal facts (codebase, team experience, contract terms) do NOT need external sources — mark them as internal assumptions instead.

## 2. Research pass (parallel)

- Fire multiple web searches at once for independent topics.
- Prefer sources in this order:
  1. **Official vendor/authority pages**: apple.com/vn, lg.com/vn (or /sg for specs), developer.apple.com, support.google.com, firebase.google.com/pricing, github.com (repo state), chinhphu.vn / vanban.chinhphu.vn (legal).
  2. **Major VN retailers** (for actual market prices): CellphoneS, GearVN, FPT Shop, Thế Giới Di Động, TopZone, Hacom, Phúc Anh, An Khang, Phong Vũ, Laptop World.
  3. Avoid blogspam/aggregators as primary sources; use them only to discover exact model names, then verify on (1)/(2).
- For spec claims, cite the vendor spec page (e.g. LG product page for USB-C 65W Power Delivery), not the retailer copy.

## 3. Capture rules (prices especially)

| Rule | Example |
|---|---|
| Record BOTH list price and sale price | "33.990.000đ → 24.990.000đ" |
| Note short-lived promotions | "giá khuyến mãi ngắn ngày, kiểm tra lại khi mua" |
| Pin the exact configuration | "24GB/512GB", not "bản cao cấp" |
| Check availability/discontinued status | e.g. "LG 24MD4KL đã ngừng sản xuất — không còn dòng 24\" 4K chính hãng" |
| Multiple retailers → cite 2–3 | "15.990.000đ (CellphoneS) [39]; 16.490.000đ (GearVN) [40]" |
| Date-stamp everything | closing line: "Giá tra ngày 15/10/2026; cần kiểm tra lại khi mua." |

## 4. Citation system

- Number sources `[1]…[n]` in order of first use; reuse the same number everywhere.
- Body text uses `[n]` only; full URLs live in appendices.
- **Appendix A — Giá thiết bị, màn hình và tài khoản**: one bullet/row per item: what it is, price, retailer + URL, `[n]` refs.
- **Appendix B — Nguồn thông tin cho các nhận định chính**: URL for each key claim (archived repo, docs page, platform guideline, reference implementation).

Example appendix line (Vietnamese, keep format):

```
- Màn hình LG 24QP750-B — 5.190.000đ: CellphoneS — https://cellphones.com.vn/man-hinh-gaming-lg-24qp750-24-inch.html; GearVN — https://gearvn.com/products/man-hinh-lg-24qp750-b-24-ips-2k-75hz-hdr10-usbc. Thông số USB-C 65W (LG): https://www.lg.com/sg/consumer-monitors/fhd-qhd/24qp750-b/
```

## 5. Common traps

- Retailer pages change price silently — always include the checked date.
- Aggregator "cheapest" numbers may come from grey-market sellers; verify the retailer is authorized (chính hãng).
- Discontinued-but-still-listed products: check the manufacturer page for "Discontinued" banners before proposing them.
- When a product doesn't exist in a size/type the user wants (e.g. 24" 4K còn bán), say so explicitly in the document and offer the nearest real option — never invent a listing.
