# Đợt 4 — Excel và dấu vết soát lại (4.3.1 → 4.4.0)

Viết ngày 16/11/2026, nhánh `claude/dot-4-a4b5e7` (gốc `origin/main` = `150e4ca`, bản đang chạy 4.3.1). Việc lấy từ mục "Đợt 4" của `tasks/ke-hoach-v4.md`, cộng các việc tồn nhỏ mà `tasks/HANDOFF.md` và `tasks/dot/dot-3/dot-3.md` ghi để lại.

## Trạng thái (cập nhật sau mỗi chặng)

- [x] 0. Tệp này được commit trước khi viết mã (`8ba546a`)
- [x] 1. Thăm dò Excel (bảng kết quả ở dưới) · [x] đo số lần gọi AI của việc Excel trên bản 4.3.1 (mốc so sánh: 6 lần gọi)
- [x] 2. Mã lõi thuần + kiểm thử máy: nhật ký, hàm nguy hiểm, kiểm toán mô hình, giải ngược
- [x] 3. Công cụ Excel mới (9 công cụ, Excel nay 29) + 15 bước Tự kiểm tra
- [x] 4. Việc tồn: câu lỗi "mở lại cửa sổ Word", lỗi 413 của Groq, bảng nghiệm thu cuối `dot-3.md`; phần model miễn phí: chưa kiểm được (OpenRouter hết lượt)
- [x] 5. Tự kiểm tra ba ứng dụng trên bản cục bộ · [x] Chạy thật các tiêu chí · [x] đo lại số lần gọi AI (6 → 3–4)
- [x] 6. Soát mã độc lập (3 cao, 17 trung bình, 16 thấp), đã sửa 24, để lại 12 có lý do, đã chạy lại
- [x] 7. Bảng nghiệm thu, HANDOFF, CLAUDE.md, bảng so sánh · [x] anh duyệt, **đã phát hành 4.4.0** · [x] kiểm bản đã phát hành · [x] dọn dẹp

**Trạng thái của khung trước khi thử (để trả lại khi xong):** đọc bằng mã lúc mở Excel: nguồn đang chọn Anthropic · Claude Opus 5.5; Groq: GPT OSS 20B; OpenRouter: ERNIE 5.5 Flash free (theo ghi chú dọn dẹp của bản 4.3.1); hai công tắc nhật ký và tô sáng: tắt. Bài thử đã chọn Haiku 5.5 (Anthropic) và **đã làm model nhớ của OpenRouter đổi sang "Nemotron 3 Nano Omni (free)"** — phải trả về ERNIE 5.5 Flash free khi dọn.

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
| T-4 | Mục "Bảng nghiệm thu" cuối `tasks/dot/dot-3/dot-3.md` hết trống (bảng tổng hợp trỏ về các bảng A, B đã có bằng chứng) | Đọc tệp |

### Chung

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| C-1 | `npm test` đạt hết (423 cũ + mới), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |
| C-2 | Mỗi công cụ Excel mới hoặc sửa có bước Tự kiểm tra, đạt hết; Word 37, Excel 23, PowerPoint 21 bước cũ không hỏng | Tự kiểm tra |
| C-3 | Soát mã độc lập đã chạy trên `git diff 150e4ca..HEAD -- src test`; phát hiện đã sửa hoặc ghi rõ lý do để lại | Báo cáo soát mã |
| C-4 | Thư mục `tasks/dot/dot-4/bang-chung/` và mọi tệp của đợt không lộ key, token, email, họ tên, tên tài khoản Office, số dư, mã tổ chức | Soát bằng mã trước mỗi commit |
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

**5. Kiểm trong Office** bằng `kiem-thu/mo-office.ps1` + `kiem-thu/cdp.mjs serve` (bản cục bộ): Tự kiểm tra Excel, Word, PowerPoint; Chạy thật D-1, D-2, D-4, D-6, D-10, T-1, T-2 (Claude Haiku 5.5 như các đợt trước; T-2 bằng Groq). Dữ liệu thử dựng ngay trong sổ tính bằng `kiem-thu/cdp-bai/dot-4-excel-du-lieu.js` (bảng tiến độ để in; mô hình tài chính cài 3 hằng số gõ cứng + 1 công thức lệch hàng; bảng theo dõi hồ sơ có dòng trùng); bài giao việc `dot-4-excel-bai.js`, công tắc `dot-4-cong-tac.js`, Word `dot-4-word-gui.js`. Bằng chứng vào `tasks/dot/dot-4/bang-chung/`.

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

Chạy 16/11/2026 trong Excel 16.0 (ExcelApi 1.9–1.20 đều có), bài `kiem-thu/cdp-bai/tham-do-4-excel.js` chạy Office.js thẳng trong khung add-in, chỉ đụng hai sheet thử rồi xóa; bằng chứng `tasks/dot/dot-4/bang-chung/01-tham-do-excel.json`. Chữ ký API đối chiếu với `@types/office-js` của kho. Đã đo, không làm trái:

| Điều thử | Kết quả | Dùng cho |
|---|---|---|
| Biểu đồ có sẵn: đổi `chartType`, `title.text`, `setData`, tiêu đề trục, vị trí chú giải | **Được.** Nguồn của từng chuỗi đọc được bằng `series.getDimensionDataSourceString("Values")`; nhãn trục (`"Categories"`) trả null | `excel_edit_chart`, `excel_list_objects` |
| Pivot có sẵn: thêm trường cột, thêm giá trị và đổi hàm (`summarizeBy`), bớt trường, lọc theo mục (`applyFilter({ manualFilter })`), `clearAllFilters`, `refresh` | **Được.** Thêm giá trị thứ hai thì trục cột có thêm mục giả "Values" (phải lọc khi liệt kê); tên trường giá trị là "Sum of …" (bớt theo tên đó); sửa dữ liệu nguồn thì Pivot KHÔNG tự đổi cho tới khi `refresh` | `excel_edit_pivot` |
| `pivotTable.getDataSourceString()` | **Được** (trả vùng nguồn) | `excel_list_objects` |
| `worksheet.pageLayout`: hướng giấy, `paperSize = "A4"`, `zoom = { horizontalFitToPages: 1 }`, `setPrintArea`, `setPrintTitleRows`, `setPrintMargins("Centimeters", …)`, căn giữa, đầu / chân trang | **Được, đọc lại đúng** (`getPrintAreaOrNullObject`, `getPrintTitleRowsOrNullObject`). `paperSize` ban đầu đọc ra null; bỏ vùng in bằng `setPrintArea("")` được; bỏ co vừa trang bằng `zoom = { scale: 100 }` | `excel_page_setup` |
| Tên vùng: `names.add(tên có dấu tiếng Việt, range, ghi chú)`, dùng trong công thức, liệt kê, xóa | **Được.** Xóa tên thì công thức đang dùng ra `#NAME?` → xóa tên là thao tác phải hỏi | `excel_named_range` |
| `range.formulasR1C1`: cùng một công thức kéo sang ô khác cho chuỗi giống hệt; hằng số trả về kiểu số | **Đúng** → so chuỗi R1C1 là đủ để tìm ô lệch và số gõ cứng | `excel_audit_model` |
| Ghi ô đầu vào rồi đọc ô phụ thuộc trong CÙNG một lần `sync` | **Ra số đã tính lại**, khoảng 38 ms mỗi vòng | `excel_what_if` |
| Chế độ tính tay (`calculationMode = "Manual"`) | Ô phụ thuộc KHÔNG đổi cho tới khi gọi `calculate` → công cụ tự gọi `application.calculate("Full")` khi sổ ở chế độ tính tay | `excel_what_if` |
| `removeDuplicates(cột, cóTiêuĐề)`, `replaceAll`, `autoFill` (`FillDefault`, `FillWeekdays`) | **Được**, trả số dòng đã xóa / số chỗ đã thay. `Range` không có `findAll` (chỉ `Worksheet` có) | `excel_clean_data` |
| `range.group("ByRows" / "ByColumns")`, `ungroup`, `hideGroupDetails` | **Được** | nhóm trong `excel_insert_delete` |
| Khóa ô công thức: bỏ khóa cả vùng dùng, khóa `getSpecialCells("Formulas")`, `protection.protect({…})` không mật khẩu | **Được.** Ghi vào ô khóa bị `AccessDenied`, ghi ô nhập vẫn được; sau khi bỏ bảo vệ, `locked` của vùng đọc ra null (lẫn lộn) | `excel_protect`; câu lỗi gợi ý "sheet đang bảo vệ" |
| `workbook.linkedWorkbooks` | Có (0 liên kết ở tệp thử); công cụ kiểm toán dò liên kết ngoài bằng mẫu trong công thức | `excel_audit_model` |
| Màu nền: ô không tô đọc ra `""`; `fill.clear()` trả về không tô | **Đúng** → phân biệt được "không màu" khi trả màu cũ | tô sáng, `clearHighlights` |

Kết luận: mọi việc của Đợt 4 làm được bằng API có sẵn, không cần phương án dự phòng. Giải ngược và bảng nhạy cảm làm bằng cách add-in tự thử từng giá trị rồi trả ô về nội dung gốc (Goal Seek và Data Table gốc của Excel không mở cho add-in).

## Bảng nghiệm thu

Kiểm 16/11/2026 trên **bản cục bộ 4.4.0** nạp vào khung add-in bằng `cdp.mjs serve` (gói `taskpane.2e7b9026.js`, mã sau khi sửa theo soát mã). Chạy thật dùng Anthropic · Claude Haiku 5.5, trừ T-2 (Groq). Bằng chứng ở `tasks/dot/dot-4/bang-chung/` (số tệp ở cột cuối). Kết quả trên bản đã phát hành: mục "Phát hành" bên dưới.

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| D-1 | Bật nhật ký, 5 thao tác ghi → đủ 5 dòng (giờ, ô, cũ, mới, lý do, yêu cầu); mặc định tắt; có cảnh báo khi bật | **Đạt** — Tự kiểm tra (5 lần ghi → 5 dòng, công thức ghi thành chữ, AI không ghi được vào sheet nhật ký). Chạy thật: bật hai công tắc bằng cú bấm trong Cài đặt (trước đó cả hai tắt; câu cảnh báo "đi theo tệp" hiện khi bật), AI sửa D6, E9, C12, B3, B4 → 5 dòng đủ 7 cột, đúng giá trị cũ / mới. Cột "lý do": Haiku không nói gì trước khi gọi công cụ nên cột này chỉ có tên thao tác; model có nói thì lời đó được chép vào (Kiểm thử máy) | `14`, `05`; `test/unit/dot4.test.ts` |
| D-2 | Hoàn tác: ô hết tô sáng, nhật ký thêm dòng hoàn tác | **Đạt** — Chạy thật: 5 ô được tô `#FFF2CC`; bấm "Hoàn tác thay đổi (4)" → 5 ô về giá trị cũ, nền về không màu, nhật ký thêm 4 dòng "Hoàn tác: …". Tự kiểm tra: ô có sẵn màu xanh thì về đúng màu xanh | `05`, `14` |
| D-3 | Đổi kiểu và tiêu đề biểu đồ có sẵn; thêm trường vào Pivot có sẵn | **Đạt** — Tự kiểm tra: cột → đường, tiêu đề mới, đọc lại đúng, Hoàn tác trả về; Pivot thêm trường cột, đổi hàm sang trung bình, lọc một mục, bớt trường, làm mới, đọc lại đúng | `14` |
| D-4 | "Chuẩn bị in bảng tiến độ": vùng in, A4 ngang, vừa một trang ngang, dòng tiêu đề lặp | **Đạt** — Tự kiểm tra: đặt đủ 4 thứ + lề + chân trang rồi đọc lại đúng, Hoàn tác về giấy dọc không vùng in. Chạy thật (2 lần gọi AI, 5 giây): A4 ngang, vừa 1 trang ngang, lặp dòng 1; AI không đặt vùng in riêng (in cả vùng dữ liệu A1:E13) vì yêu cầu không nêu vùng | `14`, `06` |
| D-5 | Mô hình cài 3 hằng số + 1 công thức lệch: tìm đủ 4; bản sạch không báo nhầm | **Đạt** — Tự kiểm tra (đủ 4 đúng ô; sửa xong thì 0), Kiểm thử máy (cả các bố cục dễ báo nhầm từ soát mã), Chạy thật: AI nêu đúng D6, E9, C12 (số gõ cứng) và E10 (lệch), kèm hai tỷ lệ gõ trong công thức | `14`, `07`; `dot4.test.ts`, `dot4-soat.test.ts` |
| D-6 | Giải ngược đạt mục tiêu trong 0,1%; không ghi thì mọi ô về giá trị gốc | **Đạt** — Tự kiểm tra: lệch 0,0000 % sau 14 vòng, ô đầu vào vẫn 0,08, ghi thử số tìm được thì ô kết quả đúng mục tiêu; ô đầu vào là công thức bị từ chối. Chạy thật: tìm B4 = 0,5909655 để B13 = 500, AI xin ghi → bấm "Bỏ qua" → B4 vẫn 0,62, B13 vẫn 429,096 | `14`, `08`, `08b` |
| D-7 | Công thức có WEBSERVICE hoặc INDIRECT: có hỏi trước khi ghi | **Đạt** — Tự kiểm tra (hai hàm có hỏi; SUM và chữ thường không hỏi); Kiểm thử máy (DDE, liên kết tệp ngoài, values dạng chuỗi JSON, ca dễ báo nhầm). Giới hạn: người dùng tắt "Luôn hỏi trước khi ghi đè…" thì hộp này cũng tắt | `14`; kiểm thử máy |
| D-8 | Tên vùng; bảng nhạy cảm 1–2 biến, trả ô về giá trị gốc | **Đạt** — Tự kiểm tra | `14` |
| D-9 | Xóa dòng trùng (có hỏi), thay hàng loạt, tự điền, nhóm, khóa ô công thức không mật khẩu; hoàn tác | **Đạt** — Tự kiểm tra (5 bước) | `14` |
| D-10 | Việc Excel 6 yêu cầu: số lần gọi AI trước / sau | **Đạt** — 4.3.1: **6 lần gọi**, 18 lệnh trong đó 2 lệnh lỗi (địa chỉ nhiều vùng), 23 giây, biểu đồ phải dựng khối phụ H1:J13. 4.4.0: **4 lần** rồi **3 lần** (hai lần chạy), 9–10 lệnh, 0 lỗi, 16 giây, biểu đồ cột vẽ thẳng từ hai vùng rời; công thức và Pivot đọc lại đúng | `02a`; `04`, `04b` |
| T-1 | Word thu nhỏ: câu gợi ý mở lại cửa sổ | **Đạt** — Chạy thật: thu nhỏ Word bằng COM → khung hiện câu gợi ý, AI nhắn lại đúng ý, 0 nhận xét; khôi phục cửa sổ, gửi lại → 1 nhận xét | `11` |
| T-2 | Groq 413: tiếng Việt, gợi ý đổi nguồn, không lộ mã tổ chức | **Đạt** — Chạy thật trong Word (Groq · GPT OSS 20B): câu tiếng Việt có số token, nút Đổi nguồn, 0 chuỗi dạng mã tổ chức trên khung; Kiểm thử máy cho "Chép chi tiết lỗi" và tên model có dấu chấm | `12`, `12b`; kiểm thử máy |
| T-3 | Model miễn phí của Đợt 3 | **Chưa kiểm được** — OpenRouter trả "hết lượt model miễn phí hôm nay" (có lại 7 giờ sáng) | `13` |
| T-4 | Bảng nghiệm thu cuối `dot-3.md` | **Đạt** | `tasks/dot/dot-3/dot-3.md` |
| C-1 | `npm test`, tsc, đóng gói | **Đạt** — 470/470 (423 cũ + 47 mới), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |
| C-2 | Tự kiểm tra ba ứng dụng | **Đạt** trên mã cuối — Excel **38/38** (23 cũ + 15 mới), Word **37/37**, PowerPoint **21/21** | `14`, `15`, `16` (lần đầu: `03`, `09`, `10`) |
| C-3 | Soát mã độc lập | **Đạt** — 235 ca thử tự viết; 3 cao, 17 trung bình, 16 thấp; sửa 24, để lại 12 (bảng dưới) | `soat-ma-doc-lap.md` |
| C-4 | Không lộ thông tin riêng | **Đạt** — soát bằng mã trước mỗi commit (mẫu key, mã tổ chức, email, tên trong hồ sơ Office): không có | lệnh soát trong nhật ký phiên |
| C-5 | Lịch sử chỉ thêm, lời nhắc cố định; nhật ký không vào lịch sử gửi model | **Đạt** — kiểm thử chặn `dot0b` vẫn đạt; lời nhắc Excel cố định (kiểm thử); tác tử soát mã xác nhận `said` / `request` chỉ đi vào ngữ cảnh của công cụ | Kiểm thử máy; `soat-ma-doc-lap.md` |

## Quyết định của anh (16/11/2026, ô tích chọn)

1. **Phát hành 4.4.0.**
2. Tô sáng ô vừa ghi: **mặc định tắt** (như nhật ký).
3. Câu hỏi trước khi ghi công thức có hàm rủi ro: **tắt theo công tắc chung** "Luôn hỏi trước khi ghi đè hoặc xóa" (không gắn cờ luôn hỏi).

## Phát hành 4.4.0 (16/11/2026, anh duyệt bằng ô tích chọn)

- Đẩy tiến thẳng `150e4ca..dcc372b` lên `main`; GitHub Actions đạt; `gh-pages` "Bản chạy thật v4.4.0 (main@dcc372b)"; Pages phục vụ `taskpane.2e7b9026.js` (trùng gói đóng cục bộ đã kiểm).
- Kiểm trên bản đã phát hành (ba khung tải từ GitHub Pages, đã dừng nạp bản cục bộ): đầu khung v4.4; Tự kiểm tra **Excel 38/38** (`17`), **Word 37/37** (`18`), **PowerPoint 21/21** (`19`); Chạy thật trong Excel (Haiku 5.5, `20`): bật hai công tắc bằng cú bấm (có câu cảnh báo), AI sửa 5 ô trong 3 lần gọi → 5 dòng nhật ký, ô tô `#FFF2CC`; bấm Hoàn tác → ô về giá trị cũ, hết màu, nhật ký thành 10 dòng; tắt lại hai công tắc.
- Dọn dẹp: xóa hội thoại thử (Excel còn 1, Word còn 4, PowerPoint còn 2 — bằng số trước đợt); model nhớ của OpenRouter trả về ERNIE 5.5 Flash free, Groq GPT OSS 20B, nguồn đang chọn Anthropic · Claude Opus 5.5; hai công tắc nhật ký và tô sáng tắt; Word, Excel, PowerPoint thử đã đóng (không lưu), cổng 9229 đã đóng.

## Mã của đợt

- **Lõi thuần:** `src/core/auditlog.ts` (dòng nhật ký, chữ ghi vào nhật ký luôn là chữ, nhận diện sheet nhật ký), `formularisk.ts` (công thức cần hỏi), `modelaudit.ts` (kiểm toán mô hình trên lưới R1C1), `goalseek.ts` (giải ngược: kẹp nghiệm rồi dây cung có chặn).
- **Office:** `src/office/excelModel.ts` — 9 công cụ: `excel_list_objects`, `excel_edit_chart`, `excel_edit_pivot`, `excel_page_setup`, `excel_named_range`, `excel_audit_model`, `excel_what_if`, `excel_clean_data`, `excel_protect`; `excel_insert_delete` thêm nhóm / bỏ nhóm; `excel_format_range` nhận nhiều vùng rời; `excel_create_chart` nhận cột nhãn + khối số nằm rời; `excel_write_range` hỏi trước công thức nguy hiểm. Excel nay **29 công cụ**.
- **Dấu vết soát lại:** `src/office/excelTrail.ts` — lớp bọc quanh bộ công cụ Excel (`getHostTools`): nhật ký, tô sáng, chặn AI ghi vào sheet nhật ký, gỡ dữ liệu nội bộ khỏi kết quả, câu gợi ý khi sheet đang bảo vệ. Hai công tắc ở Cài đặt (chỉ hiện trong Excel), mặc định tắt, kèm nút "Bỏ tô sáng".
- **Ngữ cảnh cho nhật ký:** vòng lặp đưa lời model của bước (`said`) và câu người dùng gõ (`request`) cho công cụ; không thêm gì vào lịch sử gửi model.
- **Việc tồn:** `errors.ts` (lý do `too_large`, che mã tổ chức, `officeErrorHint`), `wordTools.ts` (chèn nhận xét bị từ chối → câu gợi ý).
- **Tự kiểm tra:** `src/office/selftestExcel.ts` (15 bước); phần Excel của Tự kiểm tra nay chạy với nhật ký tắt và sheet nhật ký thử riêng, không đụng "Nhật ký AI" thật.

## Soát mã độc lập (16/11/2026)

Tác tử khác đọc toàn bộ diff `150e4ca..HEAD -- src test`, tự viết 235 ca thử. Báo cáo: `tasks/dot/dot-4/bang-chung/soat-ma-doc-lap.md`.

**Đã sửa (24), có kiểm thử ở `test/unit/dot4-soat.test.ts`:**

| # | Mức | Việc |
|---|---|---|
| 1 | Cao | Giải ngược với mục tiêu = 0 (hòa vốn, NPV = 0) trên số cỡ tỷ đồng báo "chưa hội tụ" dù đã đúng → sai số tính theo thang đầu ra |
| 2 | Cao | `values` tới dưới dạng chuỗi JSON thì công thức nguy hiểm được ghi không hỏi → quét và đếm trên đúng khối sẽ ghi |
| 3 | Cao | Câu 413 của Groq với tên model có dấu chấm không ra câu tiếng Việt |
| 4, 17, 23, 24, 25 | TB / Thấp | Công thức cần hỏi: DDE ở mọi vị trí; liên kết tệp ngoài hết báo nhầm `[2024]`, bắt `[Book1]Sheet1!` và tên vùng của tệp khác; gạch đầu dòng "- …" không còn bị coi là công thức; liên kết nội bộ không hỏi; thêm FILTERXML, REGISTER |
| 5, 6, 7, 34 | TB / Thấp | Sheet nhật ký: chặn cả lệnh không nêu sheet khi đang đứng ở sheet nhật ký; chặn tên mà sau khi làm sạch thành "Nhật ký AI"; xóa sheet nhật ký không còn bị tạo lại ngay; công cụ tổng quan không trả nội dung nhật ký |
| 8, 9 | TB | Khóa ô công thức: mở khóa cả sheet (thêm được dòng mới), kiểm điều kiện xong mới đổi khóa, báo rõ khi sheet đã có bố trí khóa riêng |
| 10, 11, 12 | TB | Tự điền ghi đè dữ liệu thì hỏi; câu hỏi nói đúng khi nào không hoàn tác được; chuỗi rỗng không còn bị hiểu là "bỏ vùng in" |
| 13, 19 | TB | Kiểm toán: hàng / cột số nhập tay xen giữa công thức và hàng năm không còn bị báo là gõ cứng; tỷ lệ sau dấu phẩy được thấy, phần trăm trong tên cột không tính |
| 14, 15, 16, 18 | TB | `min` / `max` dạng chuỗi có hiệu lực; ô đầu vào bị khóa không còn báo sai "không trả lại được"; thay hàng loạt bằng công thức nguy hiểm thì hỏi; sửa Pivot kiểm tên trường trước khi đổi |
| 20, 36 | TB / Thấp | Tên sheet có dấu phẩy không bị cắt ngang; quá 20 vùng thì báo lỗi thay vì bỏ lặng lẽ |
| 21, 22 | Thấp | `maskSecrets` không che nhầm `project_management_2026`…, che cả dạng viết hoa và có gạch nối; câu "tokens per min" |

**Để lại (12), lý do:**

| # | Việc | Lý do |
|---|---|---|
| 9 (một phần) | Khóa ô công thức thay bố trí khóa sẵn có mà không hỏi | Đã báo trong kết quả và Hoàn tác được; thêm hộp hỏi cần đọc trạng thái khóa trong bước kiểm rủi ro — làm khi anh dùng tới sheet có khóa riêng |
| 26 | Kiểm toán bỏ sót hai ô gõ đè liền nhau, ô trống giữa dãy; báo nhầm dòng "Cộng nhóm" giữa cột | Đánh đổi có chủ ý: ưu tiên không báo nhầm; kết quả là "điểm cần xem" |
| 27 | Nhật ký của thay hàng loạt, xóa trùng, tự điền, xóa vùng không có giá trị cũ / mới từng ô; định dạng nhiều vùng chỉ ghi vùng đầu | Có dòng nhật ký nêu thao tác và vùng; ghi từng ô cần mỗi công cụ trả thêm dữ liệu — đợt sau |
| 28, 29, 30 | Hoàn tác chưa đủ: tiêu đề trục biểu đồ, thứ tự trường và lọc cũ của Pivot, vùng in nhiều vùng và khổ giấy lạ | Hiếm; kết quả công cụ đã ghi rõ phần không trả lại được (nguồn dữ liệu biểu đồ) |
| 31 | Giải ngược quá 120 giây vẫn chạy nền rồi mới trả gốc; tải lại khung giữa chừng để ô ở giá trị thử | Mỗi vòng khoảng 40 ms, tối đa 60 vòng (bảng nhạy cảm tối đa 225) nên rất khó chạm 120 giây; tải lại khung giữa chừng thì không mã nào cứu được |
| 32 | Sổ ở chế độ tính tay: mỗi vòng tính lại cả sổ | Bắt buộc để đọc được kết quả; giá trị ô nhập không đổi |
| 33 | Số dạng chữ kiểu "1.000.000", "50,60" bị từ chối | Lược đồ khai kiểu số; model gửi số |
| 35 | Màu cũ của tô sáng chỉ nhớ trong phiên; chèn / xóa hàng làm lệch địa chỉ đã nhớ | Giới hạn đã ghi; Hoàn tác (đường chính) không phụ thuộc bộ nhớ này |
| 36 (một phần) | Biểu đồ hai vùng có thể để lại biểu đồ dở khi lỗi giữa chừng | Hiếm; người dùng xóa được |
| — | Hộp hỏi công thức nguy hiểm tắt theo công tắc "Luôn hỏi trước khi ghi đè…" | Cùng cơ chế với mọi cảnh báo ghi đè; muốn không tắt được thì gắn cờ `alwaysAsk` — hỏi anh nếu cần |

## Giới hạn đã biết của 4.4.0

- Nhật ký chỉ ghi thao tác của AI qua add-in (không ghi việc anh tự sửa tay); cột "lý do" là lời model nói ở bước đó, model không nói thì chỉ có tên thao tác. Sheet nhật ký nằm cuối sổ, đi theo tệp.
- Tô sáng đổi màu nền ô; Hoàn tác trả màu cũ. Giữ thay đổi mà muốn bỏ màu: nút "Bỏ tô sáng" trong Cài đặt, chỉ nhớ các ô tô trong phiên đang mở.
- Giải ngược và bảng nhạy cảm chỉ nhận ô đầu vào là số gõ trực tiếp; không tự ghi kết quả (muốn ghi thì AI dùng lệnh ghi, có hỏi vì ghi đè).
- Sửa nguồn dữ liệu biểu đồ và xóa biểu đồ không hoàn tác được. Bảo vệ sheet không mật khẩu: chống sửa nhầm, không phải bảo mật.
- Kiểm toán mô hình dò theo mẫu, cho "điểm cần xem"; chỉ xét ô nằm giữa dãy công thức.
- Chưa Chạy thật các công cụ mới với model miễn phí (OpenRouter hết lượt; Groq gói miễn phí không chạy nổi một yêu cầu của add-in).
