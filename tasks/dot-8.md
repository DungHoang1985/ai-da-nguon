# Đợt 8 — điền mẫu trong Word (bản 4.8.0, kế hoạch)

Viết ngày 18/11/2026, gốc `main@f8c5a01`, bản đang chạy 4.7.2. Anh chọn việc này ở mục D (`tasks/dot-5.md`) ngày 18/11/2026.

Trạng thái: **chờ anh duyệt kế hoạch và trả lời các câu ở mục "Chờ anh chọn"**. Chưa sửa dòng mã nào; đã thăm dò Office.js trong Word thật (desktop ẩn).

Mọi bước kiểm trong Office chạy trên desktop ẩn (`kiem-thu/mo-office.ps1`). Anh đang mở Word thì phiên dừng và hỏi, không mở bản thử.

## Mục tiêu

Anh đưa một mẫu Word (hợp đồng, tờ khai, giấy đề nghị…) cùng thông tin (gõ trong câu, tệp đính kèm, bảng Excel đang mở) → AI liệt kê mọi chỗ cần điền, điền hết **trong một lệnh**, có Track Changes, một nút Hoàn tác trả lại y như cũ; chỗ nào thiếu thông tin thì để nguyên và kê ra cho anh, **không bịa**.

Bản chính thức có "điền mẫu theo kiểu chữ của văn bản" [W]; bảng so sánh dòng 45 đang ghi "Một phần — chưa điền ô mẫu".

## Hiện trạng (đọc mã, 18/11/2026)

- Đọc được ô điền mẫu: `word_read_parts` (phần `content_controls`) trả tên, tag, loại, chữ, ô trống — `src/office/wordReview.ts:350`. Soát nhất quán đã báo "ô chưa điền" và chỗ trống "……", "____", "[●]", "ngày   tháng   năm" — `src/core/consistency.ts:526`.
- **Chưa có công cụ điền.** Muốn điền thì AI phải sửa từng đoạn bằng `word_replace_text` (mỗi chỗ một lệnh, không vào được ô điền mẫu).
- Điền bảng có sẵn: `word_fill_table` (đã có).

## Thăm dò Office.js (18/11/2026, Word trên desktop ẩn, mẫu `kiem-thu/mau-thu.mjs`)

| Loại chỗ cần điền | Cách ghi được | Track Changes | Từ chối sửa đổi trả lại y cũ | Bằng chứng |
|---|---|---|---|---|
| Ô chữ thường (PlainText), ô chữ định dạng (RichText) | `contentControl.insertText(…, "Replace")` | Có | Có (ô về trống) | `tham-do-dien-mau.json` |
| Ô đánh dấu (CheckBox) | `checkboxContentControl.isChecked` | Có (☐ xóa, ☒ thêm) | Có | như trên |
| Ô chọn hoặc nhập (ComboBox), ô ngày (DatePicker) | `insertText` trên ô **báo lỗi**; `getRange("Content").insertText` **chạy** | Có | Có | `tham-do-o-chon.json` |
| Ô chọn (DropDownList) | Không nhận chữ (cả khi tắt theo dõi); API đọc danh sách mục không có trên bản Word này (WordApi 1.9 có, 1.10 không). **Đọc danh sách từ OOXML của ô, ghi lại ô bằng `insertOoxml` với mục đã chọn: chạy** | Có (ô cũ chờ xóa, ô mới chờ thêm) | Có (đủ 8 ô như cũ) | `tham-do-o-chon-ooxml.json` |
| Ô khóa nội dung | Word từ chối (đúng) | — | — | `tham-do-dien-mau.json` |
| Chỗ trống "……", "[●]" trong đoạn thường | Tìm **đúng trọn dãy** trong đoạn (lấy từ chữ của đoạn), chọn đúng lần xuất hiện thứ k, `insertText Replace` | Có | Có | `tham-do-dien-mau.json`, `tham-do-o-chon.json` |
| Ô bảng trống | `cell.body.insertText` | Có | Có | `tham-do-dien-mau.json` |

Lưu ý đo được: tìm theo một đoạn chấm ngắn hơn dãy thật chỉ thay một phần dãy; khi đang theo dõi, `paragraph.text` vẫn chứa chữ đã xóa (đọc lại phải dùng bản hiện tại như `word_find`).

## Việc

### A. Công cụ `word_fill_form` (một công cụ, hai việc)

- `action: "list"` — liệt kê mọi chỗ cần điền, mỗi chỗ một mã ngắn (`f1`, `f2`…):
  - ô điền mẫu: tên, tag, loại, giá trị hiện tại, ô khóa hay không, danh sách mục (ô chọn, ô chọn hoặc nhập — đọc từ OOXML), định dạng ngày (ô ngày);
  - chỗ trống gõ tay: đoạn [¶n], số điều, **nhãn đứng trước** ("Họ và tên người đại diện Bên B:", "Số CCCD:"), chữ quanh chỗ trống, thứ tự trong đoạn; bỏ qua dòng ký tên (dùng lại luật của soát nhất quán);
  - ô bảng trống cạnh ô nhãn ("Diện tích (m²)" → ô bên phải).
- `action: "fill"` — nhận danh sách `{ id, value }` (ô đánh dấu: `true/false`; ô chọn: một mục trong danh sách), điền **hết trong một lần gọi**, có Track Changes, **một mục Hoàn tác** (từ chối đúng các sửa đổi lệnh này thêm vào — cơ chế Hoàn tác Word hiện có).
- Kiểm trước khi ghi: ô khóa → bỏ, báo rõ; ô chọn mà giá trị không có trong danh sách → bỏ, báo danh sách hợp lệ; ô ngày → đổi về đúng định dạng của ô (dd/MM/yyyy…); chỗ trống mà chữ quanh nó đã đổi từ lúc "list" (anh sửa tay giữa chừng) → bỏ, báo đọc lại.
- **Ghi đè ô đã có chữ: hỏi trước** (cơ chế cảnh báo ghi đè hiện có). Ô / chỗ trống còn trống: không hỏi (đúng mục đích công cụ).
- Kết quả trả: đã điền / bỏ qua kèm lý do / **còn trống** (để AI kê lại cho anh những chỗ thiếu thông tin).

### B. Kỹ năng "điền mẫu" (chỉ dặn nội dung)

- Lấy thông tin từ câu của anh, tệp đính kèm (đọc theo trang), Excel đang mở (đọc chéo). Thiếu thì **để trống và kê ra**, không suy đoán số giấy tờ, ngày, số tiền.
- Số tiền: điền cả số và chữ khớp nhau (soát nhất quán đã có phép "số lệch chữ" — chạy lại sau khi điền).
- Ngày tháng theo kiểu Việt Nam; tên riêng giữ đúng chữ hoa.
- Tự nạp khi câu có "điền mẫu", "điền vào mẫu", "điền thông tin", "điền hợp đồng"…

### C. Kiểm

| # | Tiêu chí | Cách kiểm |
|---|---|---|
| C1 | `list` trên mẫu thử tìm đủ 9 ô (6 loại) + 6 dòng chỗ trống + 3 ô bảng; không nhận nhầm dòng ký tên | Kiểm thử máy (phần nhận diện) + Tự kiểm tra Word |
| C2 | `fill` điền đủ mọi loại ô trong **một** lệnh, có Track Changes; bấm Hoàn tác → văn bản y như trước (đếm ô, chữ, sửa đổi) | Tự kiểm tra Word (bước mới) |
| C3 | Ô khóa, giá trị ngoài danh sách, chỗ trống đã bị sửa tay: bỏ và báo đúng lý do, không ghi nhầm | Kiểm thử máy + Tự kiểm tra |
| C4 | Ghi đè ô đã có chữ thì hỏi; ô trống không hỏi | Tự kiểm tra |
| C5 | Chạy thật (Claude Haiku 5.5): "điền hợp đồng đặt cọc này với thông tin sau…" → đủ, đúng, **≤ 3 lần gọi AI**; thông tin thiếu thì AI kê ra, không bịa | Chạy thật |
| C6 | Chạy thật: thông tin lấy từ tệp đính kèm (giấy tờ bên mua dạng .docx/.pdf thử) | Chạy thật |
| C7 | Soát mã độc lập cho mã mới; phát hiện cao và trung bình sửa trong đợt | Báo cáo soát mã |
| C8 | Tự kiểm tra Word, Excel, PowerPoint không lùi; kiểm thử máy đạt hết | Tự kiểm tra + `npm test` |

Thêm Chạy thật với một model Gemini (luật "công cụ mới phải chạy thật với Gemini") nếu gói miễn phí còn hạn mức hôm đó; không được thì ghi "chưa kiểm được".

### D. Ngoài phạm vi (ghi rõ để anh biết)

- Ô điền mẫu trong đầu trang / chân trang; ô lặp (repeating section); ô ảnh.
- Trường nhập kiểu cũ (legacy form field, FORMTEXT — hay gặp ở mẫu .doc cũ): Office.js không có API; nếu mẫu của anh dùng loại này thì báo, tính sau.
- "Điền theo kiểu chữ riêng" (đặt kiểu theo tên) — phần còn lại của dòng 45 bảng so sánh — để đợt khác.

## Chờ anh chọn

1. **Chỗ trống gõ tay ("……", "[●]", "____")**: làm cùng ô điền mẫu (đề xuất — mẫu hợp đồng, tờ khai ở Việt Nam phần lớn dùng dấu chấm, ít dùng ô điền mẫu), hay chỉ làm ô điền mẫu?
2. **Mẫu thật để thử**: anh có mẫu hay dùng (đã bỏ thông tin thật) muốn em thử trên đó không? Nếu không, em dùng mẫu thử tự dựng (`kiem-thu/mau-thu.mjs`, hợp đồng đặt cọc bịa).
3. **Ghi đè ô đã có chữ**: hỏi trước (đề xuất) hay cứ điền?
