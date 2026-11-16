# Đợt 4 — Excel và dấu vết soát lại (4.3.1 → 4.4.0)

Viết ngày 16/11/2026, nhánh `claude/dot-4-a4b5e7` (gốc `origin/main` = `150e4ca`, bản đang chạy 4.3.1). Việc lấy từ mục "Đợt 4" của `tasks/ke-hoach-v4.md`, cộng các việc tồn nhỏ mà `tasks/HANDOFF.md` và `tasks/dot-3.md` ghi để lại.

## Trạng thái (cập nhật sau mỗi chặng)

- [x] 0. Tệp này được commit trước khi viết mã
- [ ] 1. Thăm dò Excel (bảng kết quả ở dưới) · [ ] đo số lần gọi AI của việc Excel trên bản 4.3.1 (mốc so sánh)
- [ ] 2. Mã lõi thuần + kiểm thử máy: nhật ký, hàm nguy hiểm, kiểm toán mô hình, giải ngược
- [ ] 3. Công cụ Excel mới + bước Tự kiểm tra
- [ ] 4. Việc tồn: câu lỗi "mở lại cửa sổ Word", lỗi 413 của Groq, bảng nghiệm thu cuối `dot-3.md`
- [ ] 5. Tự kiểm tra ba ứng dụng trên bản cục bộ · [ ] Chạy thật các tiêu chí · [ ] đo lại số lần gọi AI
- [ ] 6. Soát mã độc lập, sửa, chạy lại
- [ ] 7. Bảng nghiệm thu, HANDOFF, CLAUDE.md, bảng so sánh · [ ] hỏi anh phát hành · [ ] kiểm bản đã phát hành · [ ] dọn dẹp

**Trạng thái của khung trước khi thử (để trả lại khi xong):** theo ghi chú dọn dẹp của bản 4.3.1: nguồn đang chọn Anthropic · Claude Opus 5.5; Groq: GPT OSS 20B; OpenRouter: ERNIE 5.5 Flash free. Đọc lại bằng mã lúc mở Office và ghi vào đây nếu khác.

## Tiêu chí nghiệm thu

Ba mức bằng chứng: **Tự kiểm tra** (nút trong add-in, chạy trong Office) · **Chạy thật** (giao việc cho model thật trong Office) · **Kiểm thử máy** (`npm test`). Không có bằng chứng thì không ghi Đạt.

### Bảng của kế hoạch (Đợt 4)

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| D-1 | Bật nhật ký, sau 5 thao tác ghi: sheet "Nhật ký AI" có đủ 5 dòng (giờ, ô, giá trị cũ, giá trị mới, lý do, yêu cầu của lượt). Nhật ký **mặc định tắt**; lúc bật có cảnh báo sheet nhật ký đi theo tệp | Tự kiểm tra + Chạy thật |
| D-2 | Hoàn tác thì ô hết tô sáng và nhật ký ghi thêm dòng hoàn tác | Chạy thật (+ Tự kiểm tra) |
| D-3 | Đổi kiểu và tiêu đề một biểu đồ có sẵn; thêm một trường vào Pivot có sẵn | Tự kiểm tra |
| D-4 | "Chuẩn bị in bảng tiến độ": đúng vùng in, A4 ngang, vừa một trang ngang, dòng tiêu đề lặp | Tự kiểm tra (+ Chạy thật) |
| D-5 | Mô hình thử cài 3 hằng số gõ cứng và 1 công thức lệch hàng: kiểm toán tìm đủ 4; mô hình sạch không báo nhầm | Tự kiểm tra + Kiểm thử máy |
| D-6 | Giải ngược trên mô hình thử: kết quả đạt mục tiêu trong sai số 0,1%; không ghi kết quả (hoặc bấm hủy ở hộp hỏi) thì mọi ô về giá trị gốc | Tự kiểm tra + Chạy thật |
| D-7 | Công thức có WEBSERVICE hoặc INDIRECT: có hỏi trước khi ghi | Tự kiểm tra |

### Việc thêm của đợt (cùng mục việc 3, 4, 5 của kế hoạch, kế hoạch không có dòng tiêu chí riêng)

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| D-8 | Đặt tên vùng cho ô giả định, liệt kê và xóa được; bảng nhạy cảm 1–2 biến trả đúng số và trả mọi ô về giá trị gốc | Tự kiểm tra |
| D-9 | Bảng theo dõi hồ sơ: xóa dòng trùng (có hỏi), thay hàng loạt, tự điền chuỗi ngày, nhóm hàng / cột, khóa ô công thức và chừa ô nhập (không đặt mật khẩu) — mỗi việc đọc lại đúng và hoàn tác được | Tự kiểm tra |
| D-10 | Một việc Excel nhiều thao tác (chuẩn hóa bảng + bảng tổng hợp + biểu đồ + chuẩn bị in): đo số lần gọi AI trên 4.3.1 và trên bản mới, cùng model; bản mới không nhiều hơn, kết quả đúng | Chạy thật, bảng trước / sau |

### Việc tồn gộp vào đợt

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| T-1 | Word thu nhỏ cửa sổ, chèn nhận xét bị từ chối: kết quả công cụ và khung hiện câu tiếng Việt gợi ý mở lại (khôi phục) cửa sổ Word | Kiểm thử máy (hàm dịch lỗi) + Chạy thật trong Word (thu nhỏ bằng COM) |
| T-2 | Lỗi 413 của Groq gói miễn phí: câu tiếng Việt nêu rõ yêu cầu lớn hơn hạn mức mỗi phút của gói và gợi ý đổi nguồn; **mã tổ chức không còn hiện** trên khung, trong "Chép chi tiết lỗi", trong hội thoại lưu | Kiểm thử máy + Chạy thật (Groq trong Word) |
| T-3 | Phần model miễn phí của Đợt 3 (tờ trình · Nemotron free trên mã cuối, bảng so các model miễn phí, ca ERNIE nói "đã sửa") | Chạy thật — chỉ khi OpenRouter còn lượt miễn phí trong ngày; hết lượt thì ghi "Chưa kiểm được" |
| T-4 | Mục "Bảng nghiệm thu" cuối `tasks/dot-3.md` hết trống (bảng tổng hợp trỏ về các bảng A, B đã có bằng chứng) | Đọc tệp |

### Chung

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| C-1 | `npm test` đạt hết (423 cũ + mới), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |
| C-2 | Mỗi công cụ Excel mới hoặc sửa có bước Tự kiểm tra, đạt hết; Word 37, Excel 23, PowerPoint 21 bước cũ không hỏng | Tự kiểm tra |
| C-3 | Soát mã độc lập đã chạy trên `git diff 150e4ca..HEAD -- src test`; phát hiện đã sửa hoặc ghi rõ lý do để lại | Báo cáo soát mã |
| C-4 | Thư mục `tasks/dot-4-bang-chung/` và mọi tệp của đợt không lộ key, token, email, họ tên, tên tài khoản Office, số dư, mã tổ chức | Soát bằng mã trước mỗi commit |
| C-5 | Lịch sử gửi model vẫn chỉ thêm, lời nhắc hệ thống không chứa thứ hay đổi (kiểm thử chặn `dot0b` vẫn đạt); nhật ký là việc của add-in, không vào lịch sử gửi model | Kiểm thử máy |

## Các bước, theo thứ tự

**1. Thăm dò trong Excel thật trước khi viết công cụ** (ghi bảng dưới đây; đối chiếu `@types/office-js` của kho và tài liệu Office.js hiện hành): sửa biểu đồ có sẵn (`chartType`, `title`, `setData`, đọc nguồn chuỗi); Pivot có sẵn (thêm / bớt trường, hàm tổng hợp, lọc theo mục, làm mới); `worksheet.pageLayout` (vùng in, tiêu đề lặp, hướng giấy, khổ A4, co vừa trang, lề, đầu / chân trang) đọc lại được không; `names.add`; `formulasR1C1`; `removeDuplicates`, `replaceAll`, `autoFill`; `group` / `ungroup`; `format.protection.locked` + `worksheet.protection`; ghi một ô rồi đọc ô phụ thuộc trong cùng một `sync` (cho giải ngược), chế độ tính tay. Đồng thời đo mốc: số lần gọi AI của bài D-10 trên bản 4.3.1 đang phát hành.

**2. Lõi thuần (kiểm thử máy trước):**
- `src/core/auditlog.ts`: dựng các dòng nhật ký từ một thao tác (giờ, ô, cũ, mới, lý do, yêu cầu), giới hạn độ dài, dòng hoàn tác. Giá trị bắt đầu bằng `=` ghi vào sheet nhật ký thành chữ (không để nhật ký tự chạy công thức).
- `src/core/formularisk.ts`: nhận diện hàm nguy hiểm trong các giá trị sắp ghi (WEBSERVICE, FILTERXML, INDIRECT, HYPERLINK tới địa chỉ ngoài, RTD, CALL, REGISTER.ID, liên kết DDE, tham chiếu tệp ngoài) → câu cảnh báo.
- `src/core/modelaudit.ts`: từ lưới công thức R1C1 và giá trị → hằng số gõ cứng nằm giữa dãy công thức, công thức lệch so với hai bên, tỷ lệ gõ cứng trong công thức, liên kết ra tệp ngoài. Ưu tiên không báo nhầm (dòng tổng, cột nhãn).
- `src/core/goalseek.ts`: tìm đầu vào để đầu ra đạt mục tiêu (dây cung + chia đôi), hàm đánh giá do bên gọi đưa vào nên thử được ngoài Office; lưới bảng nhạy cảm.

**3. Công cụ Excel** (tệp mới `src/office/excelModel.ts`, mỗi công cụ một bước Tự kiểm tra; gom ít công cụ để không phình lời nhắc):
- `excel_list_objects` (chỉ đọc: biểu đồ, PivotTable, bảng, tên vùng, thiết lập in của sheet)
- `excel_edit_chart` · `excel_edit_pivot` · `excel_page_setup` · `excel_named_range`
- `excel_audit_model` (chỉ đọc) · `excel_what_if` (giải ngược, bảng nhạy cảm; luôn trả ô về giá trị gốc, kể cả khi lỗi giữa chừng; không tự ghi kết quả — muốn ghi thì dùng `excel_write_range`, vốn hỏi trước khi ghi đè)
- `excel_clean_data` (xóa trùng — luôn hỏi; thay hàng loạt; tự điền chuỗi) · `excel_protect` (khóa ô công thức, chừa ô nhập, bảo vệ sheet không mật khẩu; bỏ bảo vệ) · nhóm hàng / cột thêm vào `excel_insert_delete`
- Dấu vết soát lại: lớp bọc quanh mọi công cụ ghi của Excel — bật nhật ký thì sau mỗi thao tác ghi thành công, thêm dòng vào sheet "Nhật ký AI"; bật tô sáng thì tô ô vừa ghi; Hoàn tác trả lại màu cũ và ghi dòng "Hoàn tác". AI không được ghi vào sheet nhật ký. Hai công tắc trong Cài đặt (chỉ hiện ở Excel), kèm câu cảnh báo sheet đi theo tệp.
- `excel_write_range`: hỏi trước khi ghi công thức có hàm nguy hiểm.
- Lời nhắc Excel: nhắc công cụ mới, bỏ bước chọn vùng thừa nếu đo thấy tốn lượt.

**4. Việc tồn** T-1, T-2, T-4 (mã nhỏ, có kiểm thử máy); T-3 chạy khi còn lượt.

**5. Kiểm trong Office** bằng `kiem-thu/mo-office.ps1` + `kiem-thu/cdp.mjs serve` (bản cục bộ): Tự kiểm tra Excel, Word, PowerPoint; Chạy thật D-1, D-2, D-4, D-6, D-10, T-1, T-2 (Claude Haiku 5.5 như các đợt trước; T-2 bằng Groq). Tệp thử: `kiem-thu/excel-thu.mjs` (bảng tiến độ để in; mô hình tài chính cài 3 hằng số gõ cứng + 1 công thức lệch hàng; bảng theo dõi hồ sơ có dòng trùng). Bằng chứng vào `tasks/dot-4-bang-chung/`.

**6. Soát mã độc lập** trên diff `src` + `test` (giao rõ: lớp nhật ký có ghi đè dữ liệu người dùng không, giải ngược có luôn trả giá trị gốc không, bảo vệ sheet có khóa người dùng ra ngoài không, mã tổ chức / key còn lọt đường nào không, ca báo nhầm của kiểm toán mô hình và của hàm nguy hiểm) → sửa → chạy lại.

**7. Giao:** bảng nghiệm thu, HANDOFF, CLAUDE.md, bảng so sánh (dòng 30, 31, 33, 35, 36) → hỏi anh phát hành 4.4.0 → kiểm bản tải từ GitHub Pages → dọn hội thoại thử, trả nguồn và model, đóng Office thử.

## Điểm dừng để hỏi anh (ô tích chọn)

- Trước khi đẩy lên `main` (phát hành 4.4.0).
- Tô sáng ô vừa sửa mặc định bật hay tắt (kế hoạch ghi "tắt được"; tô sáng đổi màu nền ô trong tệp của anh) — hỏi cùng lúc với câu phát hành; trong lúc chờ làm theo mặc định **tắt**, đổi là một dòng.
- Trước một cách làm lệch kế hoạch, và trước khi đóng Office đang có tài liệu không phải tệp thử.

## Việc tồn: gộp hay không

| Việc tồn (nguồn: HANDOFF, `dot-3.md`) | Xử lý | Lý do |
|---|---|---|
| Câu báo lỗi "mở lại cửa sổ Word" khi chèn nhận xét bị `AccessDenied` | **Gộp** (T-1) | Nhỏ, đã rõ nguyên nhân |
| Lỗi 413 của Groq: câu tiếng Anh, lộ mã tổ chức | **Gộp** (T-2) | Nhỏ; lộ mã tổ chức là việc phải che |
| Model miễn phí của Đợt 3 (tờ trình · Nemotron, bảng so model, ca ERNIE) | **Gộp có điều kiện** (T-3) | Chỉ chạy được khi OpenRouter còn lượt miễn phí trong ngày |
| Đo và giảm số lần gọi AI cho việc Excel nhiều thao tác | **Gộp** (D-10) | Đúng phạm vi Excel |
| "Bảng nghiệm thu" cuối `dot-3.md` còn trống | **Gộp** (T-4) | Việc giấy tờ |
| Trần số bước theo loại việc kèm câu hỏi "làm tiếp?"; công cụ "thêm nhiều slide một lệnh" | Không gộp | Thuộc lõi tác tử và PowerPoint, không phải Excel; cần đo riêng |
| Công cụ chèn ảnh vào slide; ghi chéo với hai tài liệu cùng ứng dụng | Không gộp | PowerPoint / chéo ứng dụng, lệch phạm vi |
| Soát mã độc lập cho diff Đợt 0B | Không gộp | Diff cũ, lớn; làm riêng khi còn hạn mức |
| 5 bài đối chiếu với add-in Claude chính thức | Không gộp | Cần anh ngồi trước máy |
| PDF quét và Suy nghĩ kỹ với Gemini / OpenRouter / OpenAI | Không gộp | Gemini hết tiền; không thuộc Excel |
| Bố cục thừa khi chèn biểu đồ PowerPoint lần đầu | Không gộp | Office không có API xóa bố cục (đã ghi giới hạn) |

## Bảng thăm dò Excel

(điền ở bước 1)

## Bảng nghiệm thu

(điền khi có bằng chứng; bằng chứng ở `tasks/dot-4-bang-chung/`)
