# Đợt 8 — điền mẫu trong Word (bản 4.8.0, kế hoạch)

Viết ngày 18/11/2026, gốc `main@f8c5a01`, bản đang chạy 4.7.2. Anh chọn việc này ở mục D (`tasks/dot-5.md`) ngày 18/11/2026.

Trạng thái (18/11/2026): **anh duyệt kế hoạch theo ba đề xuất** (làm cả chỗ trống gõ tay; thử trên mẫu tự dựng; ghi đè ô đã có chữ thì hỏi trước). **Đã phát hành trong bản 4.8.0** (18/11/2026). Kết quả ở cuối tệp.

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

Anh trả lời 18/11/2026: duyệt theo cả ba đề xuất.

## Đã làm (18/11/2026)

- `src/core/formfill.ts` (thuần, kiểm thử máy): nhận diện chỗ trống kèm nhãn và lần xuất hiện; đọc danh sách mục, định dạng ngày từ OOXML; **dựng lại OOXML của cả ô** với giá trị mới (`fillControlOoxml`) và lý do không nên ghi đè (`controlRefusal`).
- `src/office/wordForm.ts`: `word_list_form_fields` (chỉ đọc) và `word_fill_form` (ghi). Tách hai công cụ thay cho một công cụ hai việc như kế hoạch: ở chế độ "Hỏi trước", công cụ ghi bị hỏi cả khi chỉ liệt kê.
- Kỹ năng `dien-mau` (tự nạp theo "điền mẫu", "điền giúp", "điền hợp đồng"…), một dòng trong lời nhắc hệ thống của Word.
- Tự kiểm tra Word thêm bước "Đợt 8 · Điền mẫu".

### Thăm dò thêm khi làm (đo, không đoán)

| Điều đo | Kết quả | Bằng chứng |
|---|---|---|
| Ghi chữ vào ô điền mẫu (`insertText`) rồi Hoàn tác (từ chối sửa đổi), rồi ghi lại | Ô bị để lại 5 dấu cách thay chữ gợi ý; **Word từ chối mọi lệnh ghi sau** (`GeneralException`); ô đánh dấu lệch (w14:checked = 1 mà hiện ☐) → AI báo "không điền được" ở lần hai | `chay-that-dien-mau-haiku55-lan3.json`, `tham-do-sau-hoan-tac.json` |
| Ghi đè **cả ô** bằng `insertOoxml` rồi Hoàn tác, rồi ghi lại | Trả nguyên ô cũ, ghi lại được → mọi loại ô đi đường này | `tu-kiem-tra-word-sau-soat.json` |
| Sau Hoàn tác, nội dung ô có thể thành `<w:sdtContent/>` tự đóng | Hàm dựng phải nhận dạng này | Tự kiểm tra lần 1 trượt vòng hai, đã sửa |
| Mã ô điền mẫu do Word cấp | Có thể **âm** (`c-672413777`) | `tu-kiem-tra-word-lan1.json` |
| Thay một ô chọn để lại ô cũ "chờ xóa" cạnh ô mới, cùng tên | Đọc cả hai thì ô cũ trông như còn trống → soát nhất quán báo nhầm, AI tốn 2 lượt kiểm lại | `chay-that-dien-mau-haiku55.json` (5 lần gọi) |
| Hai chỗ "[●]" cùng đoạn điền trong một lệnh | Phải đọc lại chữ hiện tại của đoạn trước mỗi chỗ (điền từ cuối lên) | `chay-that-sau-soat.json` (6 lần gọi), kiểm ngược `tu-kiem-tra-word-kiem-nguoc.json` |

### Soát mã độc lập (C7)

Báo cáo: `tasks/dot-8-bang-chung/soat-ma-dot-8.md` — **1 cao · 7 trung bình · 16 thấp**.

| # | Mức | Đã sửa | Kiểm |
|---|---|---|---|
| 1 | Cao | Điền chỗ trống và ô bảng TRƯỚC ô điền mẫu (thay ô chứa đoạn để lại đoạn cũ chờ xóa làm lệch số đoạn); `label` bắt buộc, sai nhãn thì bỏ | Tự kiểm tra (ô chứa đoạn đặt trước chỗ trống) + kiểm thử máy |
| 2 | TB | Nhãn của ô không tên = chữ đứng trước ô trong đoạn; kiểm nhãn với mọi ô | Tự kiểm tra (ô không tên) |
| 3 | TB | Bỏ kết quả tìm nằm trong chữ đã xóa chờ duyệt; đọc lại chữ hiện tại của đoạn trước mỗi chỗ | Tự kiểm tra (lệnh điền thứ hai; hai "[●]" trong một lệnh), kiểm ngược |
| 4 | TB | Chỗ trống nằm trong chữ gợi ý của ô: không liệt kê thành chỗ trống, điền thì từ chối và chỉ mã ô | — (đường mã) |
| 5 | TB | Bỏ dấu vết sửa định dạng chờ duyệt trước khi dựng OOXML (hết hở thẻ) | Kiểm thử máy |
| 6 | TB | Ô chứa bảng, hình, công thức: không ghi đè, báo sửa tay | Kiểm thử máy |
| 7 | TB | Ô khóa không cho xóa: liệt kê là khóa, không điền | — |
| 8, 9 | Thấp | Nhận thêm "năm 20……", "…/…/20…", "Họ và tên.........."; bỏ dẫn mục lục, "(…)", "“…”", cả dòng gạch — dùng chung cho soát nhất quán | Kiểm thử máy |
| 10 | Thấp | Chuỗi tìm giữ dấu cách không ngắt | Kiểm thử máy |
| 11, 12 | Thấp | Trạng thái ô đánh dấu đọc từ Word, ký hiệu dùng phông của trạng thái; bỏ đánh dấu ô đang đánh dấu thì hỏi | Kiểm thử máy |
| 13, 15, 16, 17 | Thấp | Bỏ kiểu chữ gợi ý theo mọi ngôn ngữ; định dạng ngày lạ viết dd/MM/yyyy; ký tự điều khiển, tab; ô cấp ô bảng / hàng bảng từ chối, `<w:p/>`, `<w:sdtPr/>` | Kiểm thử máy |
| 14 | Thấp | Ô gắn dữ liệu (data binding): không ghi đè | Kiểm thử máy |
| 18 | Thấp | Chỉ ẩn ô chờ xóa khi còn ô sống cùng tên | — |
| 19, 20, 21 | Thấp | Bảng đọc theo bản hiện tại; nhãn ô bảng lấy ô có chữ gần nhất bên trái + tên cột; không điền cột đầu | Kiểm thử máy |
| 24 | Thấp | Kỹ năng: sửa số lệch chữ bằng `word_replace_text`; số và chữ do người dùng đưa lệch nhau thì hỏi lại | — |

Để lại: **#22** (dòng có "chữ ký" bỏ qua cả chỗ tên trên dòng đó — thà bỏ sót còn hơn điền vào chỗ ký); **#23** (mỗi chỗ 2 lượt đồng bộ, nạp OOXML cho ô chọn / ô ngày — mẫu 20–30 chỗ chạy dưới 2 giây, đo lại khi gặp mẫu lớn).

### Kết quả theo tiêu chí

| # | Kết quả | Bằng chứng |
|---|---|---|
| C1 | Đạt — mẫu thử: 8 ô điền mẫu (6 loại, 1 ô khóa), 8 chỗ trống, 3 ô bảng; không nhận dòng ký tên | `chay-that-*.json` (bước liệt kê), kiểm thử máy `dot8*.test.ts` |
| C2 | Đạt — Tự kiểm tra Word **39/39**: 5 loại ô + ô chứa đoạn + ô không tên + chỗ trống điền trong một lệnh, một mục Hoàn tác, hoàn tác về y cũ, điền lại vòng hai rồi hoàn tác vẫn đúng | `tu-kiem-tra-word-sau-soat.json` |
| C3 | Đạt — mã sai, mục ngoài danh sách, thiếu nhãn, ô khóa: bỏ qua có lý do | như trên |
| C4 | Đạt — ghi đè ô đã có chữ thì hỏi; ô trống không hỏi | như trên |
| C5 | Đạt — Chạy thật Haiku 5.5 trên bản cuối: **3 lần gọi AI** cả hai lần (14/14, rồi bấm Hoàn tác, điền lại 15/15); thiếu CCCD thì để trống và kê lại; bản in đúng phông, không xám | `chay-that-sau-soat-1.json`, `-2.json`, `ban-in-mau-da-dien.png` |
| C6 | Đạt — thông tin Bên B lấy từ tệp .docx đính kèm: 18/18 chỗ, 4 lần gọi AI, AI chỉ ra hai thông tin trong tệp mà mẫu không có chỗ | `chay-that-tep-dinh-kem.json` |
| C7 | Đạt — báo cáo và bảng trên | `soat-ma-dot-8.md` |
| C8 | Đạt — kiểm thử máy **609/609**, Tự kiểm tra Excel 40/40, PowerPoint 25/25 (bản cục bộ trước lượt sửa theo soát mã; lượt sửa chỉ đụng mã Word và phần thuần) | `tu-kiem-tra-excel.json`, `tu-kiem-tra-powerpoint.json` |
| Gemini | **Chưa kiểm được** — Gemini 3.8 Flash nhận bộ công cụ mới (không lỗi lược đồ), gọi đúng `word_list_form_fields`; lượt điền thì máy chủ im lặng 150 giây | `chay-that-gemini.json` |

Ghi chú riêng tư: câu trả lời của AI có lúc nhắc tên tác giả sửa đổi trong Word (tên tài khoản Office trên máy); mọi tệp bằng chứng đã che tên này trước khi commit.

Dọn dẹp: model trả về Anthropic · Claude Opus 5.5, Gemini nhớ 2.5 Flash như lúc đầu; Word, Excel, PowerPoint thử đã đóng không lưu.

Phát hành 4.8.0 (18/11/2026, anh duyệt): `main@504d1d5` → `gh-pages@40771d8`, lượt Actions 44276142725 xanh, GitHub Pages trả gói `taskpane.4bc176ad.js`. Tự kiểm tra trên bản phát hành ở desktop ẩn (đầu khung hiện v4.8): Word 39/39, Excel 40/40, PowerPoint 25/25 — `phat-hanh-480-*.json`.
