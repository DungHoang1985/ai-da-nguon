# Đợt 1A — Word: đọc đúng, sửa đúng (v3.7.0 → v4.0.0)

Lập 15/11/2026 (phiên Claude Code trên máy, nhánh `claude/dot-1a-9f0fea`, gốc `217d4c4`). Việc và tiêu chí gốc: mục "Đợt 1A" của `tasks/ke-hoach-v4.md`. Bằng chứng: `tasks/dot-1a-bang-chung/`.

## Tiêu chí nghiệm thu (nêu trước khi làm)

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| 1 | Bảng kết quả thăm dò (4 kiểu đánh số × 4 đường đọc) ghi vào HANDOFF trước khi viết công cụ | Chạy trong Word qua `kiem-thu/cdp.mjs` |
| 2 | 10 câu "Điều x.y nói gì / quy định về … nằm ở điều nào" trên hợp đồng thử 50 trang: đúng số điều từ 9/10, mỗi câu không quá 4 bước công cụ | Chạy thật |
| 3 | "Viết lại Điều 4.2…": sửa đúng đoạn; sửa hai từ cách xa nhau thì Word hiện đúng hai cặp xóa/thêm; chữ đậm giữa câu còn nguyên | Tự kiểm tra + Chạy thật |
| 4 | "Thêm một khoản sau 8.2": khoản mới mang số 8.3, khoản cũ thành 8.4, cùng kiểu chữ | Tự kiểm tra |
| 5 | Chèn hai đoạn phía trên rồi sửa theo số đoạn cũ: công cụ từ chối hoặc tự định vị lại, không sửa nhầm | Kiểm thử máy + Tự kiểm tra |
| 6 | Văn bản có 30 sửa đổi và 5 nhận xét của "đối tác": tóm tắt đủ 30 kèm số điều; xử lý cả 5 nhận xét bằng Track Changes và trả lời từng nhận xét | Chạy thật |
| 7 | Sau 3 thao tác ghi và 1 nhận xét, bấm Hoàn tác: văn bản về như trước | Tự kiểm tra |
| 8 | Bảng đối chiếu từng bài với add-in Claude chính thức | Chạy thật trên cả hai (cần anh ngồi trước máy) |
| 9 | Tự kiểm tra Word chạy trong vùng cô lập: mỗi công cụ Word có ít nhất một bước có kiểm kết quả; phần văn bản ngoài vùng thử không đổi | Tự kiểm tra |
| 10 | `npm test` đạt, `npx tsc --noEmit -p .` không lỗi, đóng gói đạt; Tự kiểm tra Excel, PowerPoint không giảm | Kiểm thử máy + Tự kiểm tra |

Giữ nguyên các quy tắc của bản 3.7: lịch sử hội thoại chỉ thêm, lời nhắc hệ thống cố định, mọi sửa đổi Word ghi bằng Track Changes, thao tác rủi ro luôn hỏi kèm lý do.

## Các bước theo thứ tự, mỗi bước kiểm bằng gì

| Bước | Việc | Kiểm bằng |
|---|---|---|
| 1 | Thăm dò đánh số, đọc khi có sửa đổi, sửa nhiều chỗ trong một đoạn, chèn kế thừa đánh số, dữ liệu cho Hoàn tác | Chạy Office.js trong khung add-in (`kiem-thu/cdp-bai/tham-do-*.js`) |
| 2 | Mục lục điều khoản: `src/core/outline.ts` (thuần) + `word_get_outline`; mọi kết quả đọc kèm số điều | Kiểm thử máy + Tự kiểm tra |
| 3 | `word_find`, `word_read_clause`; lời nhắc và kỹ năng pháp lý xem mục lục trước | Kiểm thử máy + Tự kiểm tra + Chạy thật (tiêu chí 2) |
| 4 | Chống sửa nhầm đoạn: tham số `expect`, tự định vị lại hoặc từ chối; thay thế có phạm vi (đoạn nào, lần thứ mấy); hộp xác nhận xóa hiện trích nội dung | Kiểm thử máy + Tự kiểm tra (tiêu chí 5) |
| 5 | Sửa ở mức từ: `diffWords` (thuần) + áp nhiều chỗ trong một đoạn, giữ định dạng | Kiểm thử máy + Tự kiểm tra (tiêu chí 3) |
| 6 | Đọc và sửa trên văn bản đang có sửa đổi: đọc bản hiện tại / bản gốc; danh sách sửa đổi phân trang, lọc tác giả, kèm số điều và câu trước/sau | Kiểm thử máy + Tự kiểm tra + Chạy thật (tiêu chí 6) |
| 7 | Chèn điều khoản mới kế thừa kiểu và đánh số; đổi cấp | Tự kiểm tra (tiêu chí 4) |
| 8 | Hoàn tác cho Word: từ chối các sửa đổi và gỡ nhận xét của lượt vừa rồi | Kiểm thử máy + Tự kiểm tra (tiêu chí 7) |
| 9 | Tự kiểm tra Word trong vùng cô lập, đủ công cụ | Tự kiểm tra trên văn bản đang có nội dung (tiêu chí 9) |
| 10 | Nạp bản cục bộ vào khung add-in để thử trước khi phát hành (không đẩy `main`) | `kiem-thu/cdp.mjs` |
| 11 | Hợp đồng thử 50 trang (`kiem-thu/docx-thu.mjs hop-dong`), chạy các bài Chạy thật | Chạy thật |
| 12 | Cập nhật HANDOFF, bảng so sánh; hỏi anh trước khi đẩy `main` | — |

## Kết quả bước 1 — thăm dò (15/11/2026, Word 16.0.20527, khung add-in bản 3.7.0)

Văn bản thử dựng bằng `node kiem-thu/docx-thu.mjs tham-do <tệp>`: tệp .docx viết thẳng cấu trúc (đánh số nhiều cấp gắn kiểu Heading 1–3 dạng "Điều %1." / "%1.%2." / "%3)", nhiều cấp đặt trên đoạn, số gõ tay, trường SEQ / LISTNUM / AUTONUM, sửa đổi được theo dõi của tác giả "Đối tác"). Máy hỗ trợ WordApi 1.3–1.9 và WordApiDesktop 1.1–1.5. Bằng chứng: tệp 01, 02, 03.

### 4 kiểu đánh số × 4 đường đọc

| Kiểu đánh số | 1. `Paragraph.listItemOrNullObject` (WordApi 1.3) | 2. `Range.listFormat` (WordApiDesktop 1.3) | 3. `Document.listParagraphs` (WordApiDesktop 1.4) | 4. `Document.getCrossReferenceItems` (WordApiDesktop 1.4) |
|---|---|---|---|---|
| Danh sách thường "1." | **Đọc được**: "1.", "2.", "3." + cấp + thứ tự | Đọc được, thêm giá trị số | Lỗi `GeneralException` | Trả `null` |
| Nhiều cấp **gắn kiểu tiêu đề** "Điều 1." / "1.1." / "a)" | **Đọc được đúng**: "Điều 1.", "1.1.", "1.2.", "a)", "b)", "Điều 2.", "2.1." | Đọc được như đường 1 | Lỗi | Trả `null` |
| Nhiều cấp đặt trên đoạn (kiểu Normal) "1." / "1.1" / "(a)" | **Đọc được đúng** | Đọc được | Lỗi | Trả `null` |
| Số gõ tay "Điều 7.", "7.1", "a)" | Không (không phải danh sách) — số nằm sẵn trong chữ của đoạn, nhận bằng mẫu | Không | — | — |
| Trường SEQ ("Điều { SEQ }.") | Không cần: kết quả trường nằm sẵn trong chữ của đoạn ("Điều 9. Bảo mật") | — | — | — |
| Trường LISTNUM | Không | Đọc được ("1)") | — | — |
| Trường AUTONUM | Không | Không | — | — |

Kết luận:
- **Không cần phương án dự phòng tự tính số từ cấu trúc tệp.** Dùng đường 1: nạp gộp `items/listItemOrNullObject/listString,…/level,…/siblingIndex` cho mọi đoạn trong **một** lần `sync` (4 ms với 41 đoạn). Số gõ tay và trường SEQ nhận bằng mẫu trên chữ của đoạn.
- Nhãn cấp con chỉ có phần của nó ("a)"): mã phải tự ghép với cấp cha thành "1.2(a)".
- Không đọc được: số của trường LISTNUM (chỉ đường 2 đọc được, hiếm gặp) và AUTONUM (trường đời cũ). Chấp nhận; ghi vào giới hạn.
- Đường 3 và 4 không dùng được trên máy này.

### Đọc đoạn khi có sửa đổi được theo dõi

| Lối đọc | Đoạn "là ~~24 tháng~~ 12 tháng" | Đoạn đối tác đã xóa hết | Đoạn đối tác thêm mới |
|---|---|---|---|
| `paragraph.text` (bản 3.7 đang dùng) | **Lẫn chữ đã xóa**: "là 24 tháng12 tháng" | Vẫn hiện như đoạn sống | Có |
| `paragraph.getText()` (WordApi 1.7) | Đúng bản hiện tại: "là 12 tháng" | Rỗng | Có |
| `getText({IncludeTextMarkedAsDeleted})` | Như `.text` | Có | Có |
| `getReviewedText("Current")` (1.4) | Đúng bản hiện tại | Rỗng | Có |
| `getReviewedText("Original")` | Đúng bản gốc: "là 24 tháng" | Có | Rỗng |

- `.text` **có** lẫn chữ đã xóa → đọc bản hiện tại bằng `getText()` theo từng đoạn (gọi cho cả 41 đoạn trong một lần `sync`: 0,04 giây); bản gốc bằng `getReviewedText("Original")`. Chuỗi trả về có `\r` (và ký hiệu hết ô bảng) ở cuối, phải cắt.
- `search()` nhìn văn bản **gồm cả chữ đã xóa** (tìm "24 tháng" ra 1 kết quả, "là 12 tháng" ra 0 vì bị chữ đã xóa chen giữa) → mọi phép định vị bằng `search` phải tính trên chuỗi `.text`, không tính trên bản hiện tại; kết quả tìm rơi vào chữ đã xóa phải loại.
- **Thay cả đoạn trên đoạn đang có sửa đổi của đối tác làm hỏng dấu vết**: sau khi thay, "bản gốc" thành "…24 tháng12 tháng…" (sai). → Trên đoạn đang có sửa đổi chưa duyệt chỉ được sửa mức từ, không thay cả đoạn.

### Sửa nhiều chỗ trong một đoạn khi đang theo dõi thay đổi

- 4 chỗ sửa trong một đoạn, một lần `sync`, làm từ phải sang trái (đổi "30" thứ hai → "45", đổi "0,05%" → "0,03%", chèn " đủ", xóa một cụm): Word ghi đúng 6 sửa đổi (2 cặp xóa/thêm + 1 thêm + 1 xóa), bản hiện tại đúng câu mong muốn. **Được**, không phải giảm xuống sửa theo cụm.
- Định vị bằng "lần xuất hiện thứ k" của `paragraph.search` (thứ tự theo văn bản): đúng.
- Chữ "45" thay cho "30" đang đậm: vẫn đậm; chữ thường hai bên vẫn thường.
- Sửa mức từ trên đoạn đang có sửa đổi của đối tác (đổi "ký" → "bàn giao", đổi "12" trong chữ đối tác thêm → "18"): bản hiện tại đúng, bản gốc không đổi, sửa đổi của đối tác còn nguyên.

### Chèn kế thừa đánh số, định danh đoạn, Hoàn tác

- `paragraph.insertParagraph(text, "After")` sau khoản 1.1: đoạn mới **tự mang số 1.2 và cùng kiểu**, khoản cũ thành 1.3 — đúng cả với danh sách gắn kiểu tiêu đề lẫn danh sách đặt trên đoạn. Mã 3.7 hỏng vì tự gỡ khỏi danh sách và ép kiểu Normal.
- Đổi cấp bằng `listItem.level`: 1.2 → a), kiểu tiêu đề tự đổi theo (Heading 2 → Heading 3).
- `uniqueLocalId` giữ nguyên sau khi chèn đoạn phía trên và sau khi sửa chữ trong đoạn; `getParagraphByUniqueLocalId` tìm lại đúng đoạn.
- Sửa đổi do add-in tạo mang tên tài khoản Office của người dùng; **ngày chỉ chính xác tới phút** (12:03:00 trong khi giờ máy 12:03:29) → không lọc "sửa đổi của lượt này" theo mốc thời gian được. Cách làm: chụp danh sách sửa đổi trước và sau mỗi công cụ ghi, phần chênh là của AI; Hoàn tác = từ chối đúng các sửa đổi đó. Thử từ chối 12 sửa đổi của add-in: 4 sửa đổi của "Đối tác" còn nguyên, hai đoạn thử trở về đúng trạng thái trước.
- `Document.undo(1)` (WordApiDesktop 1.4) hoàn tác được một lần sửa của add-in, nhưng lẫn với thao tác tay của người dùng trong ngăn xếp Undo → không dùng làm lối chính.

## Việc đã làm và kết quả

(ghi tiếp khi làm xong từng bước)
