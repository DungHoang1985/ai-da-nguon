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
| 7 | `npm test` đạt hết (hiện 201), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |

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

Máy hỗ trợ WordApi 1.4–1.9, WordApiDesktop 1.1–1.5, WordApiHiddenDocument 1.3–1.5. Văn bản thử: . Bằng chứng: tệp 01–04 trong .

| Thứ cần đọc | API | Kết quả |
|---|---|---|
| Chú thích cuối trang, cuối văn bản |  /  (WordApi 1.5): ,  | **Đọc được**, kèm đoạn neo. Chữ chú thích mở đầu bằng ký tự  (phải bỏ) |
| Đầu trang, chân trang |  | **Đọc được** cả chữ lẫn trường PAGE |
| Bookmark | ,  (WordApi 1.4) | **Đọc được** cả bookmark ẩn (_Ref…);  (Desktop 1.4) chỉ liệt kê bookmark hiện |
| Trường dẫn chiếu | : , ,  (WordApi 1.4–1.5) | **Đọc được** (REF _Ref7001 → "7.1") |
| Ô điền mẫu |  trả rỗng;  (WordApi 1.5) | **Đọc được khi nêu rõ loại** (PlainText…); ô chưa điền trả chữ rỗng |
| Hộp văn bản |  →  (WordApiDesktop 1.2) | **Đọc được**; chữ trong hộp không nằm trong đoạn của thân văn bản |
| Lề, khổ giấy |  (WordApiDesktop 1.3), đơn vị điểm | **Đọc được** (20 mm = 56,7 điểm) |
| Phông, cỡ chữ |  | **Đọc được** theo từng đoạn |

Ghi nhận thêm:  chứa ký tự  tại mốc chú thích, còn  thì không → khi sửa đoạn phải coi mốc này là chữ "không sống" để không xóa mất chú thích (việc phải làm trong : bỏ  khỏi chữ đưa cho AI, thêm vị trí mốc vào , không thay cả đoạn khi đoạn có mốc).

### So sánh phiên bản

| Cách | Kết quả |
|---|---|
| , đích  | Chạy được (0,5 giây): Word mở cửa sổ mới "Compare Result 1" có 45 sửa đổi; **bản đang mở y nguyên** (số đoạn, chữ, 2 sửa đổi của Đối tác). Add-in **không đọc được** tài liệu kết quả (khung gắn với tài liệu gốc) |
| , đích  | Chạy được nhưng **đổi bản đang mở và làm mất 2 sửa đổi chờ duyệt của Đối tác** (50 sửa đổi mới thế chỗ) → không dùng |
|  (tài liệu ẩn) | **Đọc được tệp kia mà không mở, không đụng bản đang mở**: đủ đoạn, số điều do Word tự đánh, bản hiện tại (), sửa đổi, chú thích, đầu trang. Tài liệu ẩn nằm lại trong Word tới khi thoát (không có lệnh đóng) |

**Cách làm đã chọn** (đúng phương án dự phòng ghi ở mục 8 của kế hoạch, cộng thêm bản redline): công cụ so sánh đọc tệp đính kèm qua tài liệu ẩn, **tự so bằng mã theo từng điều khoản** để AI tóm tắt (bản đang mở không đổi); khi người dùng muốn xem redline thì gọi  với đích là tài liệu mới. Lưu ý: đây là điểm lệch nhỏ so với giả định ban đầu (kết quả so sánh của Word không đọc được từ add-in) — **cần báo anh khi xin duyệt phát hành**.

## Trạng thái khi phiên dừng vì hết hạn mức (15/11/2026)

Đã xong: bước 1, 2; bộ dựng tệp thử (bước 3, chưa có PDF quét);  (bước 4, phần hàm thuần) — chạy thử trên tệp mẫu: **bản cài lỗi ra đúng 9/9, bản sạch và bản v2 ra 0** (chưa viết thành kiểm thử máy, chưa gắn vào công cụ Word); tra tài liệu nhà cung cấp cho bước 8, 9 (tệp 05).

Chưa làm: kiểm thử máy  (dữ liệu mẫu đã có ở ); công cụ , đọc phần ngoài thân văn bản, kiểm thể thức trang, so sánh phiên bản; xử lý mốc chú thích khi sửa đoạn; PDF quét; công tắc Suy nghĩ kỹ và giữ khối suy nghĩ; Tự kiểm tra; Chạy thật; soát mã độc lập; cập nhật tài liệu; phát hành. **Chưa có tiêu chí nào đủ bằng chứng để ghi Đạt.**
