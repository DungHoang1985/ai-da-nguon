# Đợt 1A — Word: đọc đúng, sửa đúng (v3.7.0 → v4.0.0)

Lập 15/11/2026 (phiên Claude Code trên máy, nhánh `claude/dot-1a-9f0fea`, gốc `217d4c4`). Việc và tiêu chí gốc: mục "Đợt 1A" của `tasks/ke-hoach-v4.md`. Bằng chứng: `tasks/dot/dot-1a/bang-chung/`.

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

## Kết quả Đợt 1A — bản 4.0.0 (15/11/2026 buổi tối, phiên trên máy)

**Đợt 1A đạt, trừ bài đối chiếu với add-in Claude chính thức (chưa kiểm được). Đã phát hành 4.0.0** (anh duyệt 15/11/2026 buổi tối): commit `8a6dc67` trên `main`, gh-pages có "Bản chạy thật v4.0.0 (main@8a6dc67)", GitHub Pages phục vụ `taskpane.af9e4c08.js`. Kiểm lại trên bản đã phát hành (khung tải từ GitHub Pages, không nạp bản cục bộ): đầu khung v4.0, Tự kiểm tra Word 30/30 trên hợp đồng 54 trang, hai câu hỏi Chạy thật đúng số điều trong 1 bước (tệp 15–18).

Cách kiểm trong Office: Office mở bằng lệnh kèm cổng gỡ lỗi WebView2 (`kiem-thu/mo-office.ps1`), khung add-in được điều khiển bằng `kiem-thu/cdp.mjs`. Trước khi phát hành, bản 4.0.0 được nạp vào khung bằng lệnh mới `node kiem-thu/cdp.mjs serve 9229 dist`: mọi yêu cầu tới địa chỉ add-in được trả bằng tệp trong `dist/` trên máy, không đẩy gì lên mạng; đầu khung hiện "v4.0" (ảnh 11). Đây vẫn là Word thật, model thật (Anthropic · Claude Haiku 5.5). Anh theo dõi bằng điện thoại nên không xin quyền điều khiển màn hình.

### Bảng nghiệm thu

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| 1 | Bảng kết quả thăm dò ghi vào HANDOFF trước khi viết công cụ | **Đạt** — bảng ở trên và trong `tasks/HANDOFF.md`, commit `40f7380` có trước mọi commit mã | Tệp 01, 02, 03; `git log` |
| 2 | 10 câu hỏi trên hợp đồng thử 54 trang: đúng số điều từ 9/10, mỗi câu không quá 4 bước công cụ | **Đạt** — 10/10 câu nêu đúng số điều và nội dung; 8 câu 1 bước, 1 câu 2 bước, 1 câu 3 bước. (Lần chạy đầu cũng 10/10 nhưng 2 câu tốn 5 và 7 bước vì tìm theo đúng chữ quá cứng → đã cho `word_find` tìm theo từng từ và thêm một dòng lời nhắc, rồi chạy lại) | Tệp 05 (54 trang, 165 đoạn, 168.800 ký tự), tệp 06 |
| 3 | "Viết lại Điều 4.2…": sửa đúng đoạn; hai từ cách xa nhau ra đúng hai cặp xóa/thêm; chữ đậm giữa câu còn nguyên | **Đạt** — Tự kiểm tra: sửa 2 chỗ trong một đoạn ra đúng `Deleted:30 · Added:45 · Deleted:0,05% · Added:0,03%`, chữ "45" vẫn đậm, chữ hai bên vẫn thường. Chạy thật: ba chỗ sửa trong Điều 4.2 ra `Deleted "chưa" · Added "đã" · Deleted "ba" · Added "hai" · Added " Đơn giá này không áp dụng cho diện tích sân bãi."`, câu sau sửa đúng từng chữ, hai đoạn bên cạnh không đổi, sửa đổi của đối tác và sửa đổi trước đó trong đoạn còn nguyên | Tệp 04, 10 (bước "Sửa mức từ"), tệp 08 |
| 4 | "Thêm một khoản sau 8.2": khoản mới mang số 8.3, khoản cũ thành 8.4, cùng kiểu chữ | **Đạt** — Tự kiểm tra: sau 1.2 (có điểm a) khoản mới = 1.3, khoản cũ thành 1.4, cùng kiểu; khoản con = 2.1(a); đổi cấp (a) → 2.2. Chạy thật trên hợp đồng: khoản mới = 8.3, khoản cũ thành 8.4, cùng kiểu Heading 2; AI tự tìm dẫn chiếu tới 8.3 để soát | Tệp 04, 10, tệp 09 |
| 5 | Chèn hai đoạn phía trên rồi sửa theo số đoạn cũ: từ chối hoặc tự định vị lại, không sửa nhầm | **Đạt** — Kiểm thử máy (`locateParagraph`: 6 kiểm thử). Tự kiểm tra: số đoạn cũ 171 → công cụ tự sửa đúng [¶173], đoạn ở số cũ không bị đụng; nội dung không khớp → từ chối | `test/unit/dot1a.test.ts`; tệp 04, 10 |
| 6 | 30 sửa đổi và 5 nhận xét của "đối tác": tóm tắt đủ 30 kèm số điều; xử lý cả 5 nhận xét bằng Track Changes và trả lời từng nhận xét | **Đạt** — Lượt 1 (1 bước công cụ, 23 giây): nêu đúng tổng số 30, đủ 18/18 điều có sửa đổi kèm số điều và số đoạn. Lượt 2: 5 chỗ sửa đúng yêu cầu từng nhận xét (420.000 đồng; 02 tháng; trong vòng 05 Ngày Làm Việc; 05 năm; 20 ngày), ghi thành 10 sửa đổi có theo dõi; 5/5 nhận xét có câu trả lời; 30 sửa đổi của đối tác còn nguyên; không đánh dấu đã xử lý (đúng yêu cầu) | Tệp 07 |
| 7 | Sau 3 thao tác ghi và 1 nhận xét, bấm Hoàn tác: văn bản về như trước | **Đạt** — Tự kiểm tra: 4 thao tác (viết lại đoạn, thay cụm từ, chèn khoản, thêm nhận xét) rồi hoàn tác: số sửa đổi, số nhận xét, số đoạn, chữ của vùng thử y như trước; kể cả khoản chèn sát một khoản đang chờ duyệt (Word gộp hai sửa đổi) và lần sửa bên trong khoản chờ duyệt. Chạy thật: bấm nút "Hoàn tác thay đổi (3)" trên khung sau bài viết lại Điều 4.2 → khung báo "Đã hoàn tác 3 thao tác", đoạn và số sửa đổi (40) về đúng như trước | Tệp 04, 10 (bước "Hoàn tác"), tệp 08 |
| 8 | Bảng đối chiếu từng bài với add-in Claude chính thức | **Chưa kiểm được** — cần anh ngồi trước máy mở add-in Claude chính thức (có tài khoản của anh) và chạy cùng 5 bài; bài và tệp thử đã sẵn (xem mục "Việc cần anh") | — |
| 9 | Tự kiểm tra Word trong vùng cô lập, mỗi công cụ có bước, phần ngoài vùng thử không đổi | **Đạt 30/30** trên văn bản trống và trên hợp đồng 54 trang (165 đoạn, 30 sửa đổi, 5 nhận xét ngoài vùng thử); các lần chạy trước khi sửa theo soát mã còn đạt trên văn bản thăm dò (41 đoạn, 7 sửa đổi) và trên hợp đồng đang có 41 sửa đổi. Bước cuối so lại phần ngoài vùng thử: số đoạn, chữ, kiểu, số sửa đổi, số nhận xét, chế độ theo dõi đều y nguyên. 23/23 công cụ Word có bước | Tệp 04 (hợp đồng), tệp 10 (văn bản trống) |
| 10 | `npm test`, kiểm kiểu, đóng gói; Excel, PowerPoint không giảm | **Đạt** — `npm test` 201/201 (150 cũ + 51 mới), `npx tsc --noEmit -p .` sạch, `npm run build` đạt. Tự kiểm tra bản 4.0: Excel 22/22; PowerPoint 13/14 lần đầu trên bài mới rồi 14/14 lần hai (đúng như bản 3.7.0: lỗi cũ "Biểu đồ gốc thêm một bố cục", để Đợt 3) | Tệp 12, 13, 13b |

### Việc đã làm theo 9 mục của kế hoạch

1. **Thăm dò** — xong (bảng ở trên). Bước "Bộ API Word máy hỗ trợ" nằm đầu Tự kiểm tra.
2. **Mục lục điều khoản** — `src/core/outline.ts` (thuần): số do Word tự đánh (ghép nhãn cấp con với cấp cha thành "1.2(a)"), số gõ tay ("Điều 7.", "7.1", "a)", "(ii)", khoản "1." dưới "Điều 5" → 5.1), Chương / Mục / Phụ lục, tiêu đề không đánh số; phạm vi đoạn của từng điều. Công cụ `word_get_outline`. Mọi kết quả đọc kèm số điều (`num`); tổng quan văn bản gửi cho AI có mục lục cấp cao nhất.
3. **Tìm và đọc theo điều** — `word_find` (tìm trên bản hiện tại, không phân biệt hoa thường; không thấy thì thử bỏ dấu, rồi tìm theo từng từ và xếp đoạn khớp nhất lên đầu), `word_read_clause` ("Điều 5", "5.3(a)", "khoản 2 Điều 5", "Chương II"). Lời nhắc hệ thống và 4 kỹ năng pháp lý xem mục lục trước; có kiểm thử chặn kỹ năng nhắc tên công cụ không tồn tại.
4. **Chống sửa nhầm đoạn** — tham số `expect` (mấy chữ đầu của đoạn như AI đã đọc), bắt buộc với sửa đoạn, xóa đoạn, chèn điều khoản, đổi cấp: lệch thì tự định vị lại nếu chỉ một đoạn khớp, không thì từ chối kèm ứng viên. `word_replace_text` có phạm vi (đoạn nào, lần thứ mấy); nhiều chỗ khớp mà chưa chỉ rõ thì không thay, trả danh sách. Hộp xác nhận xóa hiện trích nội dung từng đoạn. Kết quả chèn trả chỉ số mới. Định danh đoạn ổn định (`uniqueLocalId`) dùng nội bộ: số đoạn của nhận xét, sửa đổi, đoạn vừa chèn, và Hoàn tác.
5. **Sửa mức từ** — `src/core/worddiff.ts` (thuần): so khác theo từ, nhiều chỗ trong một đoạn, định vị bằng "lần xuất hiện thứ k" trên chuỗi Word dùng để tìm; áp trong một lần `sync`. `word_replace_text` cũng chỉ đổi đúng từ khác đi khi cụm cũ và mới có phần chung. Viết lại gần hết một đoạn sạch (trên 60% và từ 4 chỗ) thì thay cả đoạn cho dễ đọc. Sau khi sửa, công cụ đọc lại đoạn và báo nếu câu khác câu yêu cầu.
6. **Văn bản đang có sửa đổi** — đọc bản hiện tại (`getText`), cờ `changed` / `deleted`, `view: "original"` cho bản gốc; kết quả tìm và thay bỏ qua chữ đã xóa; trên đoạn đang có sửa đổi chờ duyệt không bao giờ thay cả đoạn. `word_list_revisions`: gộp theo đoạn, kèm số điều, câu trước/sau, phân trang theo ngân sách, lọc theo tác giả, đếm theo tác giả. `word_review_revision`: đối chiếu chữ của từng sửa đổi trước khi duyệt; duyệt tất cả của một tác giả.
7. **Chèn điều khoản** — `word_insert_clause` (kế thừa kiểu và đánh số; chèn sau toàn bộ điều neo gồm khoản con; cùng cấp / xuống cấp / lên cấp; trả số mới và số của điều kế sau), `word_change_level`. `word_insert_paragraphs`: dòng "5. Thanh toán" đứng lẻ giữ nguyên số (chỉ hai dòng "1." "2." trở lên mới thành danh sách); chèn sau đoạn chữ thường thì giữ kiểu của đoạn đó; sửa lỗi tạo gạch đầu dòng sau tiêu đề có gắn đánh số.
8. **Hoàn tác cho Word** — mỗi công cụ ghi chụp danh sách sửa đổi trước và sau; phần mới hẳn được hoàn tác bằng từ chối đúng các sửa đổi đó; phần bị Word gộp vào sửa đổi có từ trước được hoàn tác bằng cách khác (gỡ đoạn đã chèn theo định danh, đưa câu về chữ cũ). Nhận xét, câu trả lời, trạng thái đã xử lý: gỡ theo id. Chấp nhận / từ chối sửa đổi: không hoàn tác được, công cụ nói rõ.
9. **Tự kiểm tra Word trong vùng cô lập** — `src/office/selftestWord.ts`, 30 bước, vùng thử nằm giữa hai đoạn mốc đặt trước đoạn cuối của văn bản.

Ngoài kế hoạch, làm kèm vì cùng chỗ: cảnh báo rủi ro Word thêm hai loại (xóa hết chữ một đoạn; xóa một cụm ở nhiều chỗ); định dạng và nhận xét phải chỉ rõ vị trí (hết mặc nhiên dùng vùng chọn); "giai đoạn 2", "công đoạn 4" không còn thành liên kết tới đoạn; `kiem-thu/cdp.mjs serve` (nạp bản cục bộ), `kiem-thu/mo-office.ps1` (mở và đóng Office thử an toàn), `kiem-thu/docx-thu.mjs` và `kiem-thu/hop-dong-thu.mjs` (dựng tệp thử).

### Soát mã độc lập (làm trong đợt này) và việc đã sửa theo đó

Một tác tử khác đọc toàn bộ `git diff 217d4c4..HEAD -- src test` và tự chạy thử các hàm thuần: 2 phát hiện mức cao, 6 trung bình, 9 thấp, 5 điểm "chưa chắc" (báo cáo nguyên văn: `tasks/dot/dot-1a/bang-chung/soat-ma-doc-lap.md`). Phần lõi so khác được thử 200.000 cặp câu ngẫu nhiên, không sai ca nào. **Đã sửa cả 17 phát hiện**, rồi chạy lại toàn bộ bài kiểm trong Word (các tệp bằng chứng 04–13b là của lần chạy lại này):

| Phát hiện | Đã sửa thế nào | Kiểm bằng |
|---|---|---|
| Cao: duyệt tất cả sửa đổi "của An" khớp cả "Hoàng Anh" | Phải khớp đúng một tên tác giả; khớp nhiều tên thì từ chối và liệt kê; hộp xác nhận ghi tên thật và số sửa đổi | Kiểm thử máy + Tự kiểm tra |
| Cao: câu thường mở đầu bằng "Điều 3 của Hợp đồng…", "Phụ lục 01 là…" bị nhận là điều khoản, làm sai cả cây mục lục | Số gõ tay chỉ được nhận khi phần sau trông như tiêu đề (có dấu phân cách sau số, hoặc ngắn, viết hoa, không kết thúc như một câu) | Kiểm thử máy (10 ca) |
| Khoản đứng trước bị xóa có theo dõi hoặc nhảy số làm khoản sau thành khoản con | Anh em là khoản đứng trước có số nhỏ hơn, không đòi liền số | Kiểm thử máy |
| Hoàn tác đưa câu về chữ cũ có thể trỏ vào chữ đã xóa | Lấy vị trí thật của khúc đã xóa; không có vị trí mà cách gán nhập nhằng thì từ chối; đọc lại sau khi sửa | Kiểm thử máy + Tự kiểm tra |
| Hoàn tác chọn nhầm sửa đổi người dùng tự làm trùng hệt ở đoạn khác | Chọn theo đoạn mà công cụ đã đụng | Kiểm thử máy |
| `expect` chưa bắt buộc ở 5 công cụ chỉ đoạn bằng số | Bắt buộc ở mọi công cụ ghi; thiếu thì từ chối | Tự kiểm tra |
| Duyệt sửa đổi theo số không cần chữ đối chiếu; duyệt nhiều không hỏi | Bắt buộc chữ đối chiếu; quá 5 sửa đổi một lúc thì hỏi | Tự kiểm tra |
| Chèn điều khoản sau điều cuối rơi xuống dưới phần chữ ký | Điều cuối: chỉ tính tới khoản con cuối cùng; đoạn mới được đưa về đúng danh sách của điều neo | Chạy thật (8.3) + Tự kiểm tra |
| 9 mục thấp | Số đo đầu đoạn ("1.5 triệu đồng"), danh sách chữ hoa gõ tay; nhận xét quá dài không làm kẹt phân trang; giới hạn cửa sổ nạp sửa đổi; thay cụm trùng hệt thì bỏ qua; giới hạn 255 tính cả ký tự ^ nhân đôi; Hoàn tác báo khi không gỡ được; công cụ ghi chạy lần lượt; **ô bảng đọc theo bản hiện tại** (thăm dò xác nhận `table.values` lẫn chữ đã xóa: "1224 tháng") | Kiểm thử máy; tệp 14 |

Ba điểm "chưa chắc" đã thăm dò trong Word (tệp 14): định dạng chữ và đổi cấp khi đang theo dõi **có** sinh sửa đổi được theo dõi (nên Hoàn tác được; bước Tự kiểm tra "Đổi cấp" nay có hoàn tác); tìm bằng chuỗi có khoảng trắng đầu hoặc cuối ra đúng số lần; đoạn trong ô bảng trả về kèm một dấu tab ở cuối (đã xử lý). Còn lại, chưa sửa: thay cả đoạn (chỉ xảy ra khi viết lại trên 60% một đoạn sạch) có thể làm mất ảnh hoặc trường nằm trong đoạn — là sửa đổi có theo dõi nên từ chối được, nhưng chưa có cảnh báo riêng.

### Giới hạn đã biết (chưa làm, có ghi chỗ)

- Số của trường LISTNUM và AUTONUM không đọc được (hiếm gặp).
- Điều cuối của văn bản kéo tới hết văn bản (gồm cả phần chữ ký) vì không có mốc kết thúc.
- Danh sách liệt kê gõ tay "1. 2." nằm trong một điều cũng gõ tay "1. 2." không phân biệt được bằng mã (danh sách do Word đánh thì phân biệt được nhờ mã danh sách).
- `word_format_text`, `word_fill_table`, `word_delete_paragraphs`, `word_replace_selection`: Hoàn tác bằng từ chối sửa đổi; nếu sửa đổi bị gộp vào sửa đổi có từ trước thì báo không hoàn tác riêng được.
- Liên kết bấm theo số điều ("Điều 5.3") trong câu trả lời: chưa làm; AI dẫn "Điều 5.3 [¶57]" và phần [¶57] bấm được.
- Bảng có ô gộp, bảng lồng; chú thích, đầu/chân trang, bookmark, hộp văn bản: Đợt 1B.
- Soát lại hình thức bằng model sau khi sửa (bản chính thức có): chưa làm; hiện công cụ tự đọc lại câu sau khi sửa.

### Ghi nhận khi kiểm

- Bản 4.0 gửi khoảng 25.800 token đầu vào cho một câu hỏi một bước (lời nhắc + 23 công cụ Word + trạng thái văn bản có mục lục 34 điều); với Anthropic gần hết đọc từ bộ nhớ đệm từ lượt hai.
- Model (Haiku 5.5) thích dùng `word_replace_text` nhiều lần hơn `word_set_paragraph_text` với đoạn dài; cả hai lối nay đều cho dấu sửa đổi nhỏ nhất.
- PowerPoint thoát chậm sau lệnh đóng (có lần hơn 20 giây); script báo "chưa thoát hẳn" nhưng tiến trình tự thoát sau đó.
- Văn bản rất dài (192 trang, 1.501 đoạn, 120 sửa đổi): mỗi lần công cụ nạp toàn văn kèm số điều mất khoảng 2,1 giây, đọc bản hiện tại của mọi đoạn thêm 0,9 giây. Chấp nhận được; nếu sau này chậm thì chỉ đọc bản hiện tại cho các đoạn có sửa đổi.

### Dọn dẹp sau khi kiểm

Đã xóa các hội thoại thử (hai lượt: 25 và 13; lịch sử Word trở về 3 như trước), chọn lại nguồn OpenRouter, đóng Word, Excel, PowerPoint không lưu tệp thử, cổng gỡ lỗi đã đóng. Tệp thử nằm trong thư mục tạm của phiên, không vào kho. Thư mục bằng chứng đã soát: không chứa tên tài khoản Office.

## Việc cần anh

1. ~~Duyệt phát hành 4.0.0~~ — anh duyệt 15/11/2026, đã phát hành và kiểm lại.
2. **Bài đối chiếu với add-in Claude chính thức** (tiêu chí 8) — anh chọn 15/11/2026: để sau, khi anh ngồi trước máy: mở tệp hợp đồng thử (tạo bằng `node kiem-thu/hop-dong-thu.mjs <tệp>`) trong Word, mở add-in Claude, lần lượt giao 5 bài: (a) "Điều 18.2(b) nói gì?"; (b) "Lãi chậm thanh toán được quy định ở điều nào và mức bao nhiêu?"; (c) "Tóm tắt toàn bộ các sửa đổi của đối tác, liệt kê đủ theo số điều"; (d) "Xử lý cả 5 nhận xét của đối tác bằng Track Changes và trả lời từng nhận xét"; (e) "Thêm một khoản mới ngay sau Điều 8.2…". Tôi ghi bảng đạt ngang / kém / hơn.
