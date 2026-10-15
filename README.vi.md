# vn-bizdoc

**Skill cho AI agent** giúp tạo tài liệu nghiệp vụ tiếng Việt chuyên nghiệp — đề xuất, báo cáo, tờ trình gửi lãnh đạo:

**nghiên cứu có dẫn nguồn → HTML chuẩn Word → xuất DOCX / PDF (khi được yêu cầu).**

Được xây dựng và kiểm chứng trên các đề xuất thực tế trình lãnh đạo (khảo sát giá tại các nhà bán lẻ Việt Nam, nhiều vòng rà soát, xử lý sự cố xuất Word).

[English version](README.md)

---

## Skill làm được gì

### Phase A — Thu thập thông tin & dẫn nguồn
- Trích xuất danh sách "số liệu cần nguồn" từ đề bài (giá, phí, chính sách nền tảng, căn cứ pháp lý, các nhận định "vì sao nên tự làm").
- Ưu tiên nguồn chính thức (Apple, LG, Google, GitHub) và các nhà bán lẻ lớn tại Việt Nam (CellphoneS, GearVN, FPT Shop, TGDD, TopZone, Hacom…), kèm quy tắc ghi nhận: giá niêm yết vs giá bán, sản phẩm ngừng sản xuất, ghi ngày tra cứu.
- Trích dẫn đánh số `[n]` xuyên suốt; danh mục nguồn đặt ở phụ lục: **A — giá thiết bị/dịch vụ/tài khoản**, **B — nguồn cho các nhận định chính**.

### Phase B — Viết & định dạng (HTML là bản gốc)
- Một file HTML duy nhất chuẩn Word: khổ A4 · lề 2 cm · Times New Roman 12pt · bảng có style · các class `.note` / `.ref` / `.center`.
- Quy tắc trình bày tiếng Việt (dấu phẩy thập phân, ngày dd/mm/yyyy, đơn vị "ngày công", danh sách từ cấm theo từng dự án).
- **Mọi URL là liên kết `<a href>` bấm được** — giữ nguyên khi nhập vào Word.
- Khung tài liệu có **lịch sử phiên bản + changelog**; đồng bộ nhiều file (bản đầy đủ và bản trình gọn độc lập); quy trình sửa theo từng vòng góp ý.

### Phase C — Xuất file (chỉ khi người dùng yêu cầu)
- **DOCX** — ưu tiên Word automation (`SaveAs2`), kèm cơ chế dự phòng tự động đã kiểm chứng cho máy bị treo SaveAs: serialize `WordOpenXML` + **đóng gói Flat OPC** (tự sinh `[Content_Types].xml`).
- **PDF** — in bằng Edge/Chrome headless (giữ nguyên thiết kế HTML và liên kết bấm được); dự phòng bằng Word `ExportAsFixedFormat`.
- Kèm kiểm chứng tự động: `officecli validate`, đối chiếu số liên kết/bảng với HTML, dọn tiến trình Word treo.

---

## Cấu trúc repository

```
vn-bizdoc/
├── SKILL.md                        # quy trình + quy tắc vàng (điểm vào cho agent)
├── README.md                       # README tiếng Anh
├── README.vi.md                    # README tiếng Việt (file này)
├── LICENSE                         # MIT
├── references/
│   ├── research-sourcing.md        # Phase A: phương pháp thu thập, nguồn ưu tiên, mẫu trích dẫn
│   ├── format-word-html.md         # Phase B: quy tắc CSS/trình bày, versioning, đồng bộ nhiều file
│   └── word-automation-notes.md    # Phase C: lỗi Word COM, pipeline dự phòng, cách kiểm chứng
├── scripts/
│   ├── export-docx.ps1             # điều phối: SaveAs → dự phòng → kiểm chứng
│   ├── _worker-saveas.ps1          # worker nền (Word SaveAs có cơ chế phát hiện treo)
│   ├── _worker-flat.ps1            # worker nền (WordOpenXML, không dùng lệnh Save)
│   ├── pack-flat-opc.ps1           # Flat OPC XML → .docx hợp lệ
│   └── export-pdf.ps1              # in headless → PDF
└── assets/
    ├── template.html               # khung HTML + CSS sẵn dùng
    └── qa-checklist.md             # checklist kiểm tra trước khi giao
```

---

## Cài đặt

Sao chép thư mục này vào thư mục skills mà agent của bạn nhận diện:

| Phạm vi | Đường dẫn (Windows) | Đường dẫn (macOS/Linux) |
|---|---|---|
| Cấp người dùng (mọi dự án) | `C:\Users\<tên-bạn>\.agents\skills\vn-bizdoc` | `~/.agents/skills/vn-bizdoc` |
| Cấp dự án | `<repo>\.agents\skills\vn-bizdoc` | `<repo>/.agents/skills/vn-bizdoc` |

Khởi động lại phiên làm việc của agent để skill được nạp.

## Cách dùng

Yêu cầu agent viết tài liệu, ví dụ:

- "Dùng vn-bizdoc viết đề xuất mua thiết bị cho đội phát triển."
- "Dùng vn-bizdoc viết báo cáo quý gửi lãnh đạo."

Skill sẽ nghiên cứu nguồn, viết file HTML hoàn chỉnh, sau đó hỏi: *"Xuất luôn .docx/.pdf hay để rà soát xong mới xuất?"*

Xuất thủ công bất cứ lúc nào:

```powershell
# DOCX (ưu tiên Word automation; tự chuyển dự phòng khi SaveAs bị treo)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\export-docx.ps1 -Html "docs\de-xuat.html"

# PDF (Edge/Chrome headless; dự phòng bằng Word)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\export-pdf.ps1 -Html "docs\de-xuat.html"
```

Tùy chọn: `-Out <đường dẫn>` (mặc định: cùng thư mục, đổi phần mở rộng) · `-Mode saveas|flat` (nhánh DOCX) · `-Mode edge|word` (nhánh PDF) · `-SaveAsTimeoutSec N`.

Mỗi lần xuất đều in kết quả kiểm chứng:

```
Parity check -> links: html=16 docx=16 | tables: html=6 docx=6
Validation passed: no errors found.
DONE: docs\de-xuat.docx
```

## Yêu cầu hệ thống

| Yêu cầu | Dùng cho |
|---|---|
| Windows + Microsoft Word | Xuất DOCX (cả nhánh chính lẫn nhánh dự phòng) |
| PowerShell 5.1 trở lên | toàn bộ scripts |
| Microsoft Edge hoặc Google Chrome | Xuất PDF (nhánh chính) — không có thì dùng Word |
| [officecli](https://d.officecli.ai) (tùy chọn) | validate tự động file xuất theo chuẩn OpenXML |

## Cơ chế dự phòng khi xuất DOCX

Một số máy gặp lỗi Word automation: `SaveAs2` treo vô hạn, không hiện hộp thoại (đã kiểm chứng: treo với mọi thư mục đích, kể cả file trắng). Script điều phối phát hiện qua timeout trên worker nền, kill Word, rồi:

1. Mở HTML bằng Word và copy nội dung sang document gốc (liên kết và bảng được giữ nguyên — đã đối chiếu 1:1).
2. Serialize document trong bộ nhớ bằng `Document.WordOpenXML` (**không dùng lệnh Save**).
3. Đóng gói Flat OPC XML thành `.docx` hợp lệ bằng `pack-flat-opc.ps1` (Flat OPC của Word thiếu `[Content_Types].xml`; script tự sinh từ thuộc tính `pkg:contentType` của từng part).

Chi tiết đầy đủ ở [`references/word-automation-notes.md`](references/word-automation-notes.md).

## Giấy phép

MIT — xem [LICENSE](LICENSE).
