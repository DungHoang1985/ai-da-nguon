# Đợt 1B — Word: soát và tra (v4.0.0 → v4.1.0)

Lập 15/11/2026 (phiên Claude Code trên máy, nhánh `claude/dot-1b-62b1fa`, gốc `a928986` = `origin/main`). Việc và tiêu chí gốc: mục "Đợt 1B" của `tasks/ke-hoach-v4.md`. Bằng chứng: `tasks/dot-1b-bang-chung/`.

## Tiêu chí nghiệm thu (nêu trước khi làm)

Sáu dòng đầu lấy đúng bảng của kế hoạch; dòng 7 là điều kiện chung của mọi đợt.

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| 1 | Cài sẵn 3 thuật ngữ lệch, 2 dẫn chiếu hỏng, 1 chỗ nhảy số, 1 số tiền lệch chữ, 2 chỗ trống: tìm đủ 9; không báo nhầm trên bản sạch | Kiểm thử máy + Chạy thật |
| 2 | Chú thích cuối trang và chân trang có trong câu trả lời khi được hỏi | Chạy thật |
| 3 | So hai phiên bản khác nhau 10 chỗ: liệt kê đủ 10 kèm số điều; bản gốc không bị đổi khi anh chưa đồng ý | Chạy thật |
| 4 | PDF quét 3 trang: tóm tắt đúng, trích đúng 3 số liệu (Anthropic; Gemini nếu có key) | Chạy thật |
| 5 | Văn bản lề sai chuẩn: báo đúng lề nào sai | Tự kiểm tra |
| 6 | Mỗi công cụ mới có bước Tự kiểm tra, đạt hết; bước cũ không hỏng (Word 30/30 hiện có, Excel 22/22, PowerPoint 14/14) | Tự kiểm tra |
| 7 | `npm test` đạt hết (đầu đợt 201), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |

Giữ nguyên các quy tắc đã chốt: lịch sử hội thoại chỉ thêm, lời nhắc hệ thống cố định (kiểm thử chặn ở `test/unit/dot0b.test.ts`); mọi sửa đổi Word ghi bằng Track Changes; thao tác rủi ro luôn hỏi kèm lý do; bước Tự kiểm tra Word chỉ đụng vùng thử; kho công khai, không đưa key hay tên tài khoản vào tệp nào.

Ba mức bằng chứng: **Tự kiểm tra** (nút trong add-in, chạy trong Word) · **Chạy thật** (giao việc cho model thật trong Word) · **Kiểm thử máy** (`npm test`). Không có bằng chứng thì không ghi Đạt.

## Các bước theo thứ tự, mỗi bước kiểm bằng gì

| Bước | Việc | Kiểm bằng |
|---|---|---|
| 1 | Viết tệp này, commit trước mọi commit mã | `git log` |
| 2 | **Thăm dò trong Word thật** (như bước 1 của Đợt 1A): (a) `compareFromBase64` — chạy được không, kết quả hiện ở đâu, add-in đọc được không, bản gốc có bị đổi không; (b) API đọc chú thích cuối trang / cuối văn bản, đầu trang và chân trang, bookmark, trường (dẫn chiếu), ô điền mẫu (content control), hộp văn bản (shape); (c) API thể thức trang (lề, khổ giấy), phông và cỡ chữ. Ghi bảng kết quả vào tệp này rồi mới viết công cụ | Chạy Office.js trong khung add-in qua `kiem-thu/cdp.mjs` (`kiem-thu/cdp-bai/tham-do-1b-*.js`) |
| 3 | Tệp thử dựng bằng mã: hợp đồng có chú thích, đầu/chân trang, bookmark, trường dẫn chiếu, ô điền mẫu, hộp văn bản và **9 lỗi cài sẵn**; kèm bản sạch, bản "phiên bản khác" lệch 10 chỗ, công văn lề sai chuẩn, PDF quét 3 trang (ảnh, không có lớp chữ) | Mở được trong Word; kiểm thử máy đọc lại cấu trúc |
| 4 | **Soát nhất quán bằng mã** (`src/core/consistency.ts`, hàm thuần, dùng lại `src/core/outline.ts`): đánh số nhảy hoặc trùng; dẫn chiếu tới điều không tồn tại; thuật ngữ định nghĩa dùng lệch hoặc dùng mà chưa định nghĩa; một bên được gọi bằng nhiều tên; chỗ trống chưa điền; số tiền bằng số không khớp bằng chữ. Công cụ `word_check_consistency` | Kiểm thử máy (9/9 trên bản cài lỗi, 0 trên bản sạch) + Tự kiểm tra + Chạy thật (tiêu chí 1) |
| 5 | **Đọc phần ngoài thân văn bản**: chú thích cuối trang / cuối văn bản, đầu trang và chân trang, bookmark, trường dẫn chiếu, ô điền mẫu, hộp văn bản — theo đúng kết quả thăm dò; cái nào máy không đọc được thì công cụ nói rõ | Tự kiểm tra + Chạy thật (tiêu chí 2) |
| 6 | **Kiểm thể thức trang**: lề, khổ giấy, phông và cỡ chữ; so với chuẩn công văn, tờ trình (không ghi số hiệu văn bản pháp luật trong mã và kỹ năng) | Kiểm thử máy + Tự kiểm tra (tiêu chí 5) |
| 7 | **So sánh với phiên bản khác** (tệp .docx đính kèm): tóm tắt khác biệt theo điều khoản; bản đang mở không bị đổi khi người dùng chưa đồng ý. Cách làm theo kết quả thăm dò 2(a); nếu `compareFromBase64` không dùng được thì **dừng và hỏi anh** trước khi chọn cách tự so bằng mã | Kiểm thử máy + Tự kiểm tra + Chạy thật (tiêu chí 3) |
| 8 | **PDF quét**: đối chiếu tài liệu chính thức hiện hành của Anthropic, Gemini, OpenRouter về gửi tệp PDF; gửi nguyên tệp cho model đọc được; nguồn khác báo rõ không đọc được | Kiểm thử máy (dựng yêu cầu) + Chạy thật (tiêu chí 4) |
| 9 | **Công tắc "Suy nghĩ kỹ"** và tùy chọn **soát lập luận trước khi ghi** sửa đổi vào điều khoản; đối chiếu tài liệu chính thức về tham số suy nghĩ của từng nguồn; giữ khối suy nghĩ trong lịch sử sao cho lịch sử vẫn chỉ thêm và bộ nhớ đệm lời nhắc không vỡ | Kiểm thử máy (kể cả các kiểm thử chặn của Đợt 0B) + Chạy thật |
| 10 | Bước Tự kiểm tra cho từng công cụ mới; chạy Tự kiểm tra Word, Excel, PowerPoint trên bản cục bộ nạp vào khung (`cdp.mjs serve`) | Tự kiểm tra (tiêu chí 6) |
| 11 | Các bài Chạy thật (tiêu chí 1–4) trên Word thật, model thật | Chạy thật |
| 12 | Soát mã độc lập: một tác tử khác đọc `git diff a928986..HEAD -- src test`; sửa theo phát hiện; chạy lại các bài kiểm | Báo cáo trong thư mục bằng chứng |
| 13 | Cập nhật `tasks/HANDOFF.md`, `CLAUDE.md`, `docs/so-sanh-claude-chinh-thuc.md`; soát thư mục bằng chứng không lộ thông tin riêng; **hỏi anh trước khi đẩy `main`** (phát hành 4.1.0) | — |
| 14 | Sau khi anh duyệt: phát hành, chạy lại Tự kiểm tra trên bản đã phát hành, dọn dẹp (xóa hội thoại thử, trả nguồn về OpenRouter, đóng Office thử, đóng cổng gỡ lỗi) | Tự kiểm tra trên bản phát hành |

Bước không đạt: làm lại riêng bước đó tối đa 3 lần rồi dừng và báo.

## Điểm phải dừng và hỏi anh

- Trước khi đẩy lên `main` (phát hành 4.1.0).
- Trước khi chọn cách làm tốn kém hoặc lệch kế hoạch (ví dụ `compareFromBase64` không dùng được và phải tự so hai tệp bằng mã).
- Trước khi đóng Office đang có tài liệu không phải tệp thử.

## Kết quả bước 2 — thăm dò (15/11/2026, Word 16.0.20527, khung add-in bản 4.0.0)

Máy hỗ trợ WordApi 1.4–1.9, WordApiDesktop 1.1–1.5, WordApiHiddenDocument 1.3–1.5. Văn bản thử: `node kiem-thu/soat-thu.mjs tham-do <tệp>`. Bài thăm dò: `kiem-thu/cdp-bai/tham-do-1b-*.js`. Bằng chứng: tệp 01–04 trong `tasks/dot-1b-bang-chung/`.

### Đọc phần ngoài thân văn bản và thể thức trang

| Thứ cần đọc | API | Kết quả |
|---|---|---|
| Chú thích cuối trang, cuối văn bản | `body.footnotes` / `body.endnotes` (WordApi 1.5): `items/body/text`, `reference` | **Đọc được**, kèm đoạn neo. Chữ chú thích mở đầu bằng ký tự mốc U+0002 (phải bỏ) |
| Đầu trang, chân trang | `section.getHeader()` / `getFooter()` với `"Primary"`, `"FirstPage"`, `"EvenPages"` (WordApi 1.1) | **Đọc được** cả chữ lẫn trường PAGE |
| Bookmark | `range.getBookmarks(includeHidden)`, `document.getBookmarkRangeOrNullObject(name)` (WordApi 1.4) | **Đọc được** cả bookmark ẩn (`_Ref…`); `document.bookmarks` (WordApiDesktop 1.4) chỉ liệt kê bookmark hiện |
| Trường dẫn chiếu | `body.fields`: `type`, `code`, `result/text` (WordApi 1.4–1.5) | **Đọc được** (`REF _Ref7001` → "7.1") |
| Ô điền mẫu | `document.contentControls` trả rỗng; `document.getContentControls({ types: […] })` (WordApi 1.5) | **Đọc được khi nêu rõ loại** (PlainText…); ô chưa điền trả chữ rỗng |
| Hộp văn bản | `body.shapes` → `shape.body.text` (WordApiDesktop 1.2) | **Đọc được**; chữ trong hộp không nằm trong đoạn của thân văn bản |
| Lề, khổ giấy | `section.pageSetup` (WordApiDesktop 1.3), đơn vị điểm | **Đọc được** (20 mm = 56,7 điểm; khổ giấy trả "A4") |
| Phông, cỡ chữ | `paragraphs` nạp `items/font/name`, `items/font/size` | **Đọc được** theo từng đoạn |

Ghi nhận thêm: `paragraph.text` chứa ký tự U+0002 tại mốc chú thích, còn `getText()` thì không. Hệ quả cho công cụ sửa đoạn có từ bản 4.0: phải coi mốc này là chữ "không sống" (như chữ đã xóa có theo dõi) để lần sửa mức từ không xóa mất chú thích; không thay cả đoạn khi đoạn có mốc; chữ đưa cho AI không kèm ký tự này.

### So sánh phiên bản

| Cách | Kết quả |
|---|---|
| `compareFromBase64`, đích `CompareTargetNew` | Chạy được (0,5 giây): Word mở cửa sổ mới "Compare Result 1" có 45 sửa đổi; **bản đang mở y nguyên** (số đoạn, chữ, 2 sửa đổi của Đối tác). Add-in **không đọc được** tài liệu kết quả (khung gắn với tài liệu gốc) |
| `compareFromBase64`, đích `CompareTargetCurrent` | Chạy được nhưng **đổi bản đang mở và làm mất 2 sửa đổi chờ duyệt của Đối tác** (50 sửa đổi mới thế chỗ, mốc chú thích bị đánh dấu xóa) → không dùng |
| `application.createDocument(base64)` (tài liệu ẩn) | **Đọc được tệp kia mà không mở, không đụng bản đang mở**: đủ đoạn, số điều do Word tự đánh, bản hiện tại (`getText()`), sửa đổi, chú thích, đầu trang. Tài liệu ẩn nằm lại trong Word tới khi thoát (API không có lệnh đóng) |

**Cách làm đã chọn** — đúng phương án dự phòng ghi ở mục 8 của kế hoạch ("so văn bản thuần"), cộng thêm bản redline của Word: công cụ so sánh đọc tệp đính kèm qua tài liệu ẩn rồi **tự so bằng mã theo từng điều khoản** để AI tóm tắt (bản đang mở không đổi); khi người dùng muốn xem redline thì gọi `compareFromBase64` với đích là tài liệu mới. Điểm lệch so với giả định ban đầu (kết quả so sánh của Word không đọc được từ add-in) được nêu lại khi xin duyệt phát hành.

## Kết quả Đợt 1B — bản 4.1.0 (15/11/2026 đêm, phiên trên máy)

**Trạng thái: mã xong, kiểm trên bản cục bộ xong, CHƯA phát hành** (chờ anh duyệt đẩy `main`). Mọi kết quả "Tự kiểm tra" và "Chạy thật" dưới đây là trên bản 4.1.0 đóng gói cục bộ nạp vào khung add-in bằng `node kiem-thu/cdp.mjs serve 9229 dist` (đầu khung hiện "v4.1"), Word / Excel / PowerPoint thật, model thật (Anthropic · Claude Haiku 5.5), và là của **lần chạy lại sau khi sửa theo soát mã độc lập**. Anh theo dõi bằng điện thoại nên không xin quyền điều khiển màn hình.

### Bảng nghiệm thu

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| 1 | Cài sẵn 3 thuật ngữ lệch, 2 dẫn chiếu hỏng, 1 chỗ nhảy số, 1 số tiền lệch chữ, 2 chỗ trống: tìm đủ 9; không báo nhầm trên bản sạch | **Đạt** — Kiểm thử máy: bản cài lỗi ra đúng 9 (thuật ngữ 3, dẫn chiếu 2, đánh số 1, số và chữ 1, chỗ trống 2), đúng đoạn đã cài; bản sạch và phiên bản khác ra 0. Chạy thật: trên bản cài lỗi công cụ trả "9 điểm cần xem: đánh số 1, dẫn chiếu 2, thuật ngữ 3, chỗ trống 2, số và chữ 1" và câu trả lời của AI nêu đủ cả 9 kèm số điều, số đoạn; trên bản sạch công cụ trả "không thấy điểm nào". Văn bản không đổi ở cả hai lần (số đoạn, số sửa đổi, mã băm chữ trước và sau như nhau). Lưu ý: ngoài 9 điểm của mã, AI còn nêu thêm các điểm do tự đọc nội dung (thiếu địa chỉ các bên, sửa đổi và nhận xét của đối tác chưa duyệt…) — đó là phần rà soát của kỹ năng "kiểm tra trước ký", không phải báo nhầm của phép soát | `test/unit/dot1b.test.ts` (5 kiểm thử "bài nghiệm thu"); tệp 07b (bản cài lỗi), 07a (bản sạch) |
| 2 | Chú thích cuối trang và chân trang có trong câu trả lời khi được hỏi | **Đạt** — hỏi "Điều 4.3 quy định gì… có chú thích nào… chân trang và đầu trang ghi gì": AI trả đúng chú thích số 2 ("Mức Phí Dịch Vụ năm 2027 do ban quản lý Tòa Nhà thông báo ngày 01/12/2026"), đầu trang ("Hợp đồng thuê văn phòng số 12/2026/HĐT-VP"), chân trang ("Bản dự thảo lần 3 ngày 10/11/2026 — Lưu hành nội bộ — Trang …") và cả hộp văn bản; 2 bước công cụ | Tệp 08 |
| 3 | So hai phiên bản khác nhau 10 chỗ: liệt kê đủ 10 kèm số điều; bản gốc không bị đổi khi anh chưa đồng ý | **Đạt** — đính kèm phiên bản khác (.docx), 1 bước công cụ: "10 chỗ khác (8 sửa, 1 chỉ ở bản đang mở, 1 chỉ ở phiên bản khác)"; AI liệt kê đủ 10 kèm số điều (3.1, 4.1, 4.2, 5.1, 6.2, 7.3, 8.2, khoản mới 9.4, khoản 11.3 bị bỏ, Phụ lục 01 mục 2.1) và nêu 11.4 thành 11.3. Văn bản đang mở không đổi: 85 đoạn, 2 sửa đổi, 1 nhận xét, mã băm chữ y nguyên. Kiểm thử máy khớp đúng 10 chỗ của bộ dựng tệp | Tệp 09; `test/unit/dot1b.test.ts` ("so phiên bản") |
| 4 | PDF quét 3 trang: tóm tắt đúng, trích đúng 3 số liệu (Anthropic; Gemini nếu có key) | **Đạt với Anthropic** — tệp 3 trang toàn ảnh (0 ký tự lớp chữ, chữ nghiêng, có nhiễu): tóm tắt đúng nội dung biên bản và trả đúng cả ba số: 48.750.000.000 đồng · 12.860 m² · 2.437.500.000 đồng. **Gemini: Chưa kiểm được** — tài khoản Gemini của anh báo hết tiền ("Your prepayment credits are depleted"); mã gửi tệp cho Gemini mới có Kiểm thử máy | Tệp 10 (Anthropic), 10b (Gemini báo hết tiền); `test/unit/dot1b-llm.test.ts` |
| 5 | Văn bản lề sai chuẩn: báo đúng lề nào sai | **Đạt** — Tự kiểm tra trên công văn thử (lề trên 15 mm, trái 20 mm): bước "Thể thức trang" báo đúng "Lề trên 15 mm (chuẩn 20–25 mm); Lề trái 20 mm (chuẩn 30–35 mm)", khớp với số đo đọc trực tiếp; trên hợp đồng lề đúng chuẩn thì báo "bốn lề và khổ giấy đúng chuẩn". Chạy thật thêm: AI nêu đúng 3 chỗ lệch (lề trên, lề trái, một đoạn phông Arial) | Tệp 13 (Tự kiểm tra), 14 (Chạy thật), 06 |
| 6 | Mỗi công cụ mới có bước Tự kiểm tra, đạt hết; bước cũ không hỏng | **Đạt** — Word **34/34** (30 bước cũ + 4 bước mới cho 4 công cụ mới) trên hợp đồng thử Đợt 1B, trên công văn và trên hợp đồng 54 trang của Đợt 1A (165 đoạn, 30 sửa đổi, 5 nhận xét ngoài vùng thử y nguyên); Excel **22/22**; PowerPoint 13/14 lần đầu trên bài mới rồi **14/14** lần hai (đúng lỗi cũ "Biểu đồ gốc thêm một bố cục", có từ bản 3.6, để Đợt 3) | Tệp 06, 13, 15 (Word); 16 (Excel); 17a, 17b (PowerPoint) |
| 7 | `npm test` đạt hết, `npx tsc --noEmit -p .` sạch, `npm run build` đạt | **Đạt** — 291/291 (201 cũ + 90 mới), kiểm kiểu sạch, đóng gói đạt | Chạy trong phiên; `test/unit/dot1b.test.ts`, `test/unit/dot1b-llm.test.ts` |

Ngoài bảng (việc 6 của kế hoạch, không có dòng tiêu chí riêng):

| Việc | Kết quả | Bằng chứng |
|---|---|---|
| Công tắc "Suy nghĩ kỹ" + giữ khối suy nghĩ trong lịch sử | **Đạt (Chạy thật, Anthropic)** — bật công tắc, bài 3 bước công cụ rồi hai lượt hỏi tiếp trong cùng hội thoại: mọi yêu cầu trả 200, có `thinking: adaptive` kèm `drop_block`, mức `high`, tiêu đề beta; số khối suy nghĩ gửi lại tăng dần 0 → 1 → 2 → 3 → 5; lượt hỏi tiếp đọc 22.512/24.259 token từ bộ nhớ đệm (lịch sử vẫn chỉ thêm, bộ nhớ đệm không vỡ). Gemini, OpenRouter, OpenAI: mới có Kiểm thử máy (dựng yêu cầu đúng tài liệu) — **Chưa kiểm được** bằng Chạy thật | Tệp 11a, 11b, 11c |
| Soát lập luận trước khi ghi sửa đổi vào điều khoản | **Đạt (Chạy thật)** — bật tùy chọn, yêu cầu sửa Điều 9.2: khung hiện "Soát lập luận trước khi ghi: … — đạt" rồi mới tới bước ghi; sửa đổi ghi bằng Track Changes (2 → 4 sửa đổi); bấm Hoàn tác về lại 2. Ca bị chặn (lượt soát trả "CHƯA ĐẠT") mới có Kiểm thử máy | Tệp 12; `test/unit/dot1b-llm.test.ts` |

### Việc đã làm theo 6 mục của kế hoạch

1. **Đọc phần ngoài thân văn bản** — công cụ `word_read_parts`: chú thích cuối trang và cuối văn bản (kèm đoạn neo và số điều), đầu trang và chân trang theo từng phần (kèm tên trường tự động như số trang), bookmark (cả bookmark ẩn khi yêu cầu), trường dẫn chiếu (mã và kết quả, đánh dấu trường hỏng), ô điền mẫu (đánh dấu ô còn trống), hộp văn bản. Ngoài ra chú thích của một đoạn đi kèm đoạn đó trong mọi kết quả đọc (`notes`), tổng quan văn bản có số chú thích. **Sửa một lỗi tiềm ẩn của bản 4.0**: `paragraph.text` chứa ký tự mốc chú thích nên viết lại một đoạn có chú thích có thể xóa mất chú thích; nay mốc được coi là chữ "không sống", đoạn có chú thích không bao giờ bị thay cả đoạn (có bước Tự kiểm tra: sửa hai lần, lần hai khi đoạn đã có chữ xóa chờ duyệt, chú thích còn nguyên).
2. **Soát nhất quán bằng mã** — `src/core/consistency.ts` (thuần, dùng lại `outline.ts`) + công cụ `word_check_consistency`: đánh số nhảy hoặc trùng (kể cả điểm a, b, c… theo bảng chữ cái tiếng Việt, khoản bị xóa có theo dõi); dẫn chiếu tới điều, khoản, điểm, phụ lục không có ("khoản 6 Điều 7", "khoản 3 Điều này", trường dẫn chiếu hỏng; bỏ qua dẫn chiếu tới luật và văn bản khác); thuật ngữ định nghĩa viết lệch hoặc viết hoa như thuật ngữ mà chưa định nghĩa; một bên nhiều tên; chỗ trống (kể cả ô điền mẫu chưa điền); số bằng số lệch bằng chữ (đọc được số tiếng Việt). Không gọi model: 1.500 đoạn xong dưới 3 giây.
3. **So sánh phiên bản** — `src/core/compare.ts` + công cụ `word_compare_versions`: đọc tệp .docx đính kèm bằng tài liệu ẩn của Word (đủ số điều do Word tự đánh), so theo đoạn, gắn số điều, nêu từng chữ đổi, khoản chỉ có ở một bản, điều khoản đổi số. Không sửa văn bản đang mở. `open_redline: true` mở bản redline của chính Word trong **cửa sổ mới** (`compareFromBase64`, đích là tài liệu mới).
4. **Thể thức trang** — `src/core/pageformat.ts` + công cụ `word_check_page_format`: khổ giấy, hướng giấy, bốn lề, phông, cỡ chữ; chuẩn mặc định là thể thức văn bản hành chính, có thể truyền chuẩn khác.
5. **PDF quét** — tệp không có lớp chữ được giữ nguyên và gửi cho model tự đọc: Anthropic (khối `document`), Gemini (`inlineData`), OpenRouter và OpenAI (phần `file`; OpenRouter chỉ rõ để model tự đọc, không tự chuyển sang OCR tính phí). Nguồn khác, hoặc model đã biết là chỉ đọc chữ: tệp không gửi đi, khung báo rõ và AI cũng được báo để nói lại với anh. Trần 15 MB mỗi tệp và khoảng 18 MB cho cả hội thoại. PDF lẫn trang quét: ghi rõ trang nào chưa đọc.
6. **Suy nghĩ kỹ và soát lập luận** — công tắc ở dấu + cạnh ô nhập và trong Cài đặt; tham số theo từng nguồn và đời model (`src/core/agent/thinking.ts`). Khối suy nghĩ của Anthropic (và `reasoning_details` của OpenRouter) được **giữ trong lịch sử và gửi lại nguyên vẹn, đúng thứ tự**; lịch sử vẫn chỉ thêm, lời nhắc hệ thống không đổi; bật tắt công tắc không đụng phần được lưu đệm. Nhà cung cấp từ chối tham số hoặc khối suy nghĩ (400) thì gửi lại một lần ở chế độ thường. Tùy chọn "Soát lập luận trước khi ghi" (Cài đặt, mặc định tắt): trước mỗi lần viết lại, thay chữ, chèn điều khoản, một lượt model riêng đối chiếu đề xuất với yêu cầu và chữ hiện tại; thấy vấn đề thì chưa ghi, trả lý do cho AI sửa đề xuất (tối đa chặn 2 lần mỗi lượt).

Kèm theo: bốn công cụ mới có nhãn tiếng Việt và hiện một dòng kết quả ngay trên khung ("9 điểm cần xem…", "10 chỗ khác…"); số mục trong phụ lục được ghi rõ "Phụ lục 01, mục 2.1"; lời nhắc hệ thống Word thêm ba dòng hướng dẫn dùng công cụ mới; `kiem-thu/soat-thu.mjs` (hợp đồng sạch / cài 9 lỗi / phiên bản khác / công văn), `kiem-thu/pdf-quet-thu.mjs` (PDF quét), các bài `kiem-thu/cdp-bai/word-1b-*.js`; `mo-office.ps1` nhận tài liệu ẩn do công cụ so sánh tạo.

### Soát mã độc lập và việc đã sửa theo đó

Một tác tử khác đọc toàn bộ `git diff a928986..HEAD -- src test` và tự chạy thử các hàm thuần: **0 phát hiện mức cao, 6 trung bình, 8 thấp**, 6 điểm "chưa chắc" (báo cáo nguyên văn: `tasks/dot-1b-bang-chung/soat-ma-doc-lap.md`). Không thấy đường làm vỡ bộ nhớ đệm, vòng gửi lại vô hạn, hay công cụ đọc ghi vào văn bản. **Đã sửa cả 14**, thêm kiểm thử hồi quy, rồi chạy lại mọi bài kiểm trong Office:

| Phát hiện | Đã sửa thế nào |
|---|---|
| M1 — thuật ngữ định nghĩa kiểu “Hợp đồng”, “Ngày làm việc” làm mọi chỗ viết thường bị báo lệch | Thuật ngữ chỉ viết hoa chữ đầu thì dạng toàn chữ thường coi là danh từ chung, không báo; thuật ngữ viết hoa từng chữ vẫn báo |
| M2 — dẫn chiếu tới luật viết tắt (BLDS, NĐ, TT…) và tới hợp đồng gốc trong phụ lục sửa đổi bị báo hỏng | Nhận tên viết tắt; nhận "Hợp đồng đã ký / gốc / thuê … ngày …" là văn bản khác; "của Hợp đồng thuê này" vẫn là văn bản đang soát |
| M3 — `1/2 (một phần hai)`, `khoản 5 (một số trường hợp)`, `hai mươi năm (20)` bị báo số lệch chữ | Bỏ qua phân số; phần chữ sau con số phải là đơn vị; "năm" đứng cuối có thể là đơn vị ở cả hai chiều |
| M4 — PDF quét nằm lại lịch sử không có trần tổng; model chỉ đọc chữ vẫn bị gửi tệp | Trần tổng cho cả hội thoại (từ chối ngay ở khung, nhắc mở hội thoại mới); model đã biết không nhận ảnh thì thay tệp bằng dòng ghi chú |
| M5 — câu trả lời Anthropic có nhiều khối chữ xen khối suy nghĩ bị gộp chữ khi gửi lại | Ghi và gửi lại từng khối chữ đúng vị trí, đúng nguyên văn |
| M6 — bước Tự kiểm tra so sánh để tệp thử nằm lại trong kho tệp đính kèm | Gỡ tệp thử trong `finally`; kho xóa khi mở hội thoại mới, tệp chờ gửi được đăng ký lại lúc gửi |
| T1–T8 | "pháp luật", "quyết định" đứng trước không còn che dẫn chiếu hỏng; dòng ký tên chỉ miễn soát chỗ trống khi ngắn; dựng biểu thức thuật ngữ một lần, trần 400 thuật ngữ; "một năm rưỡi" không báo nhầm; `word_read_parts` có phân trang; bấm Dừng lúc đang soát lập luận có sự kiện kết thúc; chỉ gửi lại "chế độ thường" với nguồn có gửi tham số suy nghĩ; PDF lẫn trang quét ghi rõ trang chưa đọc |

Điểm "chưa chắc" số 1 (vị trí mốc chú thích khi đoạn có cả chữ xóa chờ duyệt) đã đo bằng bước Tự kiểm tra mới: sửa đúng, chú thích còn nguyên. Các điểm còn lại ghi ở "Giới hạn đã biết".

Lỗi tự tìm khi chạy trong Word (đã sửa): khoản gõ tay nằm cách xa điều của nó (sau một tiêu đề khác) không được so số với khoản đứng trước — lộ ra ở lần Tự kiểm tra đầu (32/34).

### Giới hạn đã biết

- Phép soát nhất quán dò theo mẫu: kết quả là điểm cần xem. Thuật ngữ lệch **về chữ** (không phải về viết hoa) chỉ bắt được khi cụm lạ viết hoa và trùng từ đầu với một thuật ngữ đã định nghĩa ("Phí Quản Lý" ↔ "Phí Dịch Vụ"). Tên bên chỉ soát khi văn bản có định nghĩa tên bên. Số bằng chữ chỉ đọc tiếng Việt.
- So sánh phiên bản so theo chữ của đoạn (không so định dạng, bảng so theo từng ô như đoạn); mỗi lần gọi để lại một tài liệu ẩn trong Word tới khi thoát Word (Office.js không có lệnh đóng).
- Chưa Chạy thật: PDF quét và Suy nghĩ kỹ với Gemini (tài khoản hết tiền), OpenRouter, OpenAI; ca lượt soát lập luận trả "CHƯA ĐẠT"; model Claude đời cũ (nhóm dùng `budget_tokens`).
- Tệp .docx đính kèm vẫn được chép chữ vào tin nhắn (tới 120.000 ký tự) ngay cả khi chỉ để so sánh: tốn token với hợp đồng dài.
- Bảng có ô gộp, bảng lồng; liên kết bấm theo số điều; soát hình thức bằng model sau khi sửa: chưa làm (như bản 4.0).
- Khi thử PDF với Gemini, model đang nhớ của nguồn Gemini trên máy anh đã được chọn lại thành "Gemini 2.5 Flash" (không biết trước đó là model nào vì không đọc kho lưu của add-in).

### Dọn dẹp sau khi kiểm (trước khi xin duyệt)

Đã tắt hai tùy chọn thử, xóa 16 hội thoại thử (lịch sử Word về 3 như trước), chọn lại nguồn OpenRouter, đóng Word, Excel, PowerPoint không lưu tệp thử, cổng gỡ lỗi đã đóng (tệp 18). Tệp thử nằm trong thư mục tạm của phiên, không vào kho. Thư mục bằng chứng đã soát: không có tên tài khoản Office, key hay email.

## Việc cần anh

1. **Duyệt phát hành 4.1.0** (đẩy nhánh này lên `main`; GitHub Actions tự kiểm thử, đóng gói, đưa lên `gh-pages`). Sau khi phát hành tôi chạy lại Tự kiểm tra trên bản đã phát hành và dọn dẹp.
2. Biết để quyết: cách so sánh phiên bản là **tự so bằng mã + mở redline của Word ở cửa sổ mới khi cần** (lý do ở mục thăm dò); tài khoản Gemini hết tiền nên chưa thử được PDF quét với Gemini.
3. Bài đối chiếu với add-in Claude chính thức của Đợt 1A vẫn để khi anh ngồi trước máy.
