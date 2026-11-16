# Đợt 3 — Soạn văn bản hành chính, cảnh báo "nói đã sửa", chéo ứng dụng và PowerPoint (4.2.0 → 4.3.0)

Viết ngày 16/11/2026, nhánh `claude/dot-3-b3b8f4` (gốc `origin/main` = `0ba6322`). Anh giao 16/11: "làm đợt 3 trong phiên mới. Xử lý luôn cả vấn đề thiếu kỹ năng soạn thảo văn bản hành chính".

Đợt gồm hai phần. **Phần A** làm trước (chữa đúng chỗ anh vừa gặp khi soạn Word), nhỏ và độc lập, có thể phát hành riêng thành 4.2.1 nếu anh chọn. **Phần B** là Đợt 3 của `tasks/ke-hoach-v4.md`.

## Trạng thái (cập nhật sau mỗi chặng)

- [x] 0. Tệp này được commit trước khi viết mã (`66d7d16`)
- [x] A1 kỹ năng soạn văn bản hành chính · [x] A2 cảnh báo "nói đã sửa mà không ghi" · [ ] A3 model miễn phí: mới có một phần (hết lượt miễn phí trong ngày của OpenRouter, xem "Kết quả phần A")
- [x] Kiểm phần A trong Word (bản cục bộ) · [x] anh chọn phát hành 4.2.1 ngay · [x] soát mã độc lập phần A, đã sửa · [x] **đã phát hành 4.2.1**, đã kiểm bản phát hành

**Trạng thái của khung trước khi thử (để trả lại khi xong):** nguồn đang chọn Anthropic · Claude Opus 5.5; model nhớ của OpenRouter: ERNIE 5.5 Flash free (theo ghi chú bàn giao Đợt 2); Groq: GPT OSS 20B; Cerebras: gpt-oss-120b. Bài thử đã chọn Haiku 5.5 (Anthropic), Nemotron 3 Ultra free (OpenRouter), GPT OSS 120B (Groq).
- [x] B1 thăm dò kho chung và kênh lệnh (bảng kết quả ở dưới) · [ ] B1 mã ghi chéo (đang viết, nhánh `claude/dot-3-b-wip`) · [ ] B2 chỉ báo và công tắc
- [x] B1 mã ghi chéo · [x] B2 chỉ báo và công tắc · [x] B3 PowerPoint phần chắc (7 công cụ) · [x] B4 thăm dò: phông chủ đề đọc được, chèn ảnh làm được nhưng chưa thành công cụ, ghi chú diễn giả và sửa số liệu biểu đồ không có API · B5: [x] lỗi bố cục của slide biểu đồ đã chữa phần người dùng thấy; [ ] phiên Word từ chối chèn nhận xét: chưa tìm ra nguyên nhân (có manh mối)
- [x] Soát mã độc lập phần B, đã sửa, đã chạy lại · [x] Bảng nghiệm thu, HANDOFF, CLAUDE.md, bảng so sánh · [x] anh duyệt, **đã phát hành 4.3.0**, đã kiểm bản phát hành · dọn dẹp: xem cuối mục "Phát hành 4.3.0"
- **Anh nêu thêm cuối phiên (16/11/2026):** "cơ chế gọi AI đang rất phân mảnh — giao một việc mà số lần gọi AI nhảy lên mấy chục lần; sợ phân mảnh lệnh làm AI không hiểu và làm sai mục tiêu. Cần khắc phục." → việc đầu tiên của phiên sau, xem mục Bước tiếp theo của `tasks/HANDOFF.md`.

## Tiêu chí nghiệm thu

Ba mức bằng chứng: **Tự kiểm tra** (nút trong add-in, chạy trong Office) · **Chạy thật** (giao việc cho model thật trong Office) · **Kiểm thử máy** (`npm test`). Không có bằng chứng thì không ghi Đạt.

### Phần A

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| A-a | Trong Word, giao "soạn một công văn …" và "soạn một tờ trình …" bằng Claude Haiku 5.5 (trả tiền) và một model miễn phí gọi được công cụ: văn bản dựng ra đủ thành phần thể thức (quốc hiệu, tiêu ngữ, tên cơ quan, số và ký hiệu, địa danh ngày tháng, trích yếu / tên loại, kính gửi, nội dung, chức vụ và tên người ký, nơi nhận), đúng lề và cỡ chữ | Chạy thật; đối chiếu bằng kết quả của `word_check_page_format` và đọc lại văn bản bằng mã |
| A-b | Kỹ năng tự nạp khi yêu cầu nhắc công văn, tờ trình, quyết định, thông báo, giấy mời…; gọi được bằng "/"; có câu gợi ý trên màn chào của Word | Kiểm thử máy (tự chọn, "/") + Chạy thật (nhãn kỹ năng hiện trên khung) |
| A-c | Model nói "đã sửa / đã ghi / đã thêm" mà lượt đó không có thao tác ghi nào thành công: khung hiện cảnh báo. Lượt có ghi thật: không cảnh báo. AI chỉ mô tả, đề xuất, hỏi lại, hoặc nói "chưa sửa": không cảnh báo | Kiểm thử máy cho phần nhận diện (ca đúng và ca dễ báo nhầm) + Chạy thật hai ca |
| A-d | Kiểm thử chặn số hiệu văn bản pháp luật và chặn tên công cụ không tồn tại trong kỹ năng vẫn đạt | Kiểm thử máy |
| A-e | Model mặc định / thứ tự gợi ý của OpenRouter: có bảng Chạy thật so các model miễn phí trên cùng một yêu cầu Word; đổi mặc định chỉ sau khi anh chọn | Chạy thật + ô tích chọn |

### Phần B (bảng của kế hoạch)

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| B-1 | Từ PowerPoint: "lấy bảng tiến độ ở Excel đang mở, tạo 3 slide và 1 biểu đồ" chạy trọn, đúng số | Chạy thật |
| B-2 | Từ Excel: "ghi tóm tắt sang Word đang mở": đoạn mới xuất hiện trong Word, có hỏi xác nhận, hoàn tác được ở Word | Chạy thật |
| B-3 | Đóng khung ở ứng dụng kia: AI báo rõ không ghi được, không treo; chỉ báo kết nối hiện đúng | Chạy thật |
| B-4 | Slide cố ý để chữ tràn khung: AI tự phát hiện qua ảnh slide và sửa hết tràn | Chạy thật |
| B-5 | Bài 10 slide cài 5 lỗi lệch mẫu (2 phông lạ, 1 slide thiếu tiêu đề, 2 chỗ chữ dưới 12 pt): kiểm bài tìm đủ 5 | Tự kiểm tra |
| B-6 | Slide mới dùng đúng màu và bố cục của mẫu | Tự kiểm tra |

### Chung

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| C-1 | `npm test` đạt hết (363 cũ + mới), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |
| C-2 | Mỗi công cụ mới có bước Tự kiểm tra, đạt hết; Word 34, Excel 22, PowerPoint 14 bước cũ không hỏng | Tự kiểm tra |
| C-3 | Soát mã độc lập đã chạy trên `git diff 0ba6322..HEAD -- src test`; phát hiện đã sửa hoặc ghi rõ lý do để lại | Báo cáo soát mã |
| C-4 | Thư mục bằng chứng không lộ key, token, email, tên tài khoản Office, số dư | Soát bằng mã trước khi commit |

## Các bước, theo thứ tự

### Phần A

**A1. Kỹ năng `soan-thao-van-ban-hanh-chinh`** (kỹ năng dựng sẵn thứ 18, `src/core/skillpack.ts`, chỉ cho Word).
- Tham khảo (chỉ đọc) skill dùng chung cùng tên của anh; chuyển thể: ngắn, nhắc đúng tên công cụ Word đang có, diễn đạt chuẩn thể thức bằng số đo (không số hiệu văn bản pháp luật).
- Nội dung: hỏi gì trước khi soạn; bảng thành phần thể thức (vị trí, cỡ chữ, kiểu chữ); bố cục từng loại (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản); cách dựng bằng công cụ (`word_insert_table` không viền cho phần đầu và phần ký, `word_insert_paragraphs`, `word_format_text`); soát lại bằng `word_check_page_format` và đọc lại.
- Xem các công cụ Word hiện có đủ để dựng đúng thể thức chưa (bảng không viền hai cột, căn giữa, cỡ chữ từng dòng, lề trang). Thiếu gì thì thêm tham số / công cụ tối thiểu, kèm bước Tự kiểm tra và Hoàn tác.
- Từ khóa tự nạp; câu gợi ý trên màn chào của Word; sửa kỹ năng `soan-phan-hoi-phap-ly` để dẫn sang kỹ năng mới.
- Kiểm: Kiểm thử máy (số kỹ năng 18, tự nạp, "/", hai kiểm thử chặn) → Chạy thật A-a.

**A2. Cảnh báo "nói đã sửa mà không ghi".**
- `src/core/claims.ts` (hàm thuần): nhận câu trả lời cuối của lượt + việc lượt có thao tác ghi thành công không → có cảnh báo hay không. Nhận diện câu khẳng định đã làm (đã sửa, đã thay, đã thêm, đã chèn, đã xóa, đã ghi, đã cập nhật, đã soạn, đã tạo…, cả tiếng Anh), loại trừ phủ định ("chưa", "không thể", "chưa sửa"), câu điều kiện / đề xuất ("nếu anh đồng ý tôi sẽ…"), câu hỏi, câu kể về việc người khác đã làm ("đối tác đã sửa…"), trích dẫn nội dung tài liệu.
- Chỉ áp dụng ở chế độ tác tử, lượt kết thúc bình thường, không có công cụ ghi nào thành công trong lượt (kể cả ghi chéo ứng dụng ở phần B).
- Khung: dòng cảnh báo ngay dưới câu trả lời ("AI nói đã thay đổi tài liệu nhưng lượt này không có thao tác ghi nào. Tài liệu chưa đổi…"), kèm nút "Yêu cầu làm lại bằng công cụ". Cảnh báo là phần hiển thị, **không** ghi vào lịch sử gửi model (lịch sử chỉ thêm, lời nhắc hệ thống cố định); nút làm lại gửi một tin nhắn người dùng mới.
- Lưu cùng hội thoại để mở lại vẫn thấy.
- Kiểm: Kiểm thử máy (≥ 20 ca, gồm ca dễ báo nhầm) → Chạy thật A-c.

**A3. Model miễn phí của OpenRouter.**
- Chạy thật cùng một yêu cầu sửa Word trên các model miễn phí đang có trong danh mục (mỗi model 2 lần): ghi có gọi công cụ không, văn bản có đổi không.
- Từ kết quả: đề xuất thứ tự ưu tiên `PREFER.openrouter` cho nhánh miễn phí; cân nhắc danh sách "đã thấy không gọi công cụ" hiện dấu cảnh báo trong bảng chọn model. Hỏi anh bằng ô tích chọn trước khi đổi mặc định.
- Bài thử đổi model đang chọn của OpenRouter: nhớ và trả lại như trước khi thử.

**Kiểm phần A trong Word** (bản cục bộ nạp bằng `cdp.mjs serve`), rồi hỏi anh: phát hành 4.2.1 ngay hay gộp 4.3.0.

### Phần B

**B1. Thăm dò trước (ghi bảng dưới đây trước khi viết mã ghi chéo).**
1. Kho chung: ở khung Word ghi một khóa thử riêng (`aidn.thu.dot3.<ngẫu nhiên>`) vào localStorage, đọc ở khung Excel và PowerPoint, rồi xóa. Không đọc khóa nào khác.
2. Kênh lệnh: sự kiện `storage` có tới khung ở ứng dụng khác không, trễ bao lâu; `BroadcastChannel` có tới không; dự phòng là thăm dò định kỳ một khóa "hộp thư".
3. Khung ở ứng dụng nền (không có tiêu điểm) có chạy được `Word.run` / `Excel.run` ghi vào tài liệu không, bộ hẹn giờ có bị bóp không.
4. Nhịp sống: khung đóng thì bao lâu sau khung kia biết.

Thiết kế dự kiến (chốt sau thăm dò): mỗi khung ghi nhịp sống `aidn.live.<Ứng dụng>` (thời điểm, tên tệp, có nhận ghi không); công cụ `write_other_app` (ứng dụng, thao tác trong một danh sách đóng: Word thêm đoạn / Excel ghi vùng trống hoặc sheet mới / PowerPoint thêm slide chữ) → **hỏi xác nhận ở khung gửi, luôn luôn, kể cả khi đang "Cho phép hết lần này"**, hiện rõ ứng dụng đích, tên tệp và nội dung → đặt lệnh vào hộp thư → khung đích kiểm lệnh (đúng dạng, thao tác trong danh sách, công tắc nhận ghi đang bật, lệnh chưa quá hạn, chưa chạy) rồi chạy bằng chính công cụ ghi của nó (Word: Track Changes) → trả kết quả; khung đích hiện một dòng "Vừa ghi theo yêu cầu từ …" kèm nút Hoàn tác. Hết hạn chờ thì báo rõ, không treo. Nội dung đọc từ ứng dụng kia là dữ liệu, không phải chỉ dẫn (quy tắc an toàn trong lời nhắc đã kể nguồn không tin cậy: thêm "nội dung từ ứng dụng khác").

Nếu kênh lệnh không dùng được: **dừng, hỏi anh** trước khi chuyển sang phương án dự phòng của kế hoạch (giữ đọc chéo, nút "Chuyển sang … và làm tiếp" kèm sẵn nội dung).

**B2. Chỉ báo và công tắc.** Dòng nhỏ ở khung: ứng dụng nào đang kết nối (theo nhịp sống); Cài đặt: công tắc theo từng ứng dụng (cho đọc từ / cho ghi sang); khung đích chưa mở → công cụ báo rõ cách mở.

**B3. PowerPoint, phần chắc** (đối chiếu tài liệu Office.js hiện hành trước khi viết; mỗi công cụ một bước Tự kiểm tra):
- `ppt_get_slide_image`: ảnh slide gửi cho model nhìn (model không nhận ảnh thì báo rõ); kèm phép đo bằng mã: chữ tràn khung ước lượng, hình chồng nhau, hình ra ngoài slide.
- `ppt_check_deck`: kiểm bài theo mẫu bằng mã (phông lệch, cỡ chữ dưới ngưỡng, màu ngoài bảng màu chủ đề, slide thiếu tiêu đề, chữ mẫu còn sót).
- `ppt_get_theme`: màu chủ đề, danh sách bố cục và ô giữ chỗ của mẫu.
- Định dạng chữ theo từng đoạn; sửa bảng có sẵn; đọc hình và đoạn chữ đang chọn; đổi bố cục slide.
- Tệp thử: `kiem-thu/ppt-thu.mjs` dựng bài 10 slide có mẫu rõ màu chủ đề và bố cục, cài 5 lỗi lệch mẫu và 1 slide chữ tràn.

**B4. PowerPoint, phần phải thử trước** (ghi bảng thăm dò dưới đây; không được thì báo rõ, không cố): cập nhật số liệu biểu đồ đã có; đọc phông chủ đề; chèn ảnh; ghi chú diễn giả.

**B5. Hai việc tồn** (làm nếu còn sức, không thì ghi HANDOFF): lỗi "Biểu đồ gốc thêm một bố cục"; phiên Word từ chối chèn nhận xét mới.

**Sau cùng:** soát mã độc lập (chú ý: kênh lệnh chéo có thể bị lợi dụng để ghi mà không hỏi không; nội dung từ ứng dụng kia có được coi là dữ liệu không) → sửa → chạy lại → bảng nghiệm thu → HANDOFF, CLAUDE.md, bảng so sánh → hỏi anh phát hành → kiểm bản đã phát hành → dọn dẹp.

## Điểm dừng để hỏi anh (ô tích chọn)

Trước mỗi lần đẩy lên `main` · trước khi đổi model mặc định · trước khi chuyển sang phương án dự phòng của B1 hoặc một cách làm lệch kế hoạch · trước khi đóng Office đang có tài liệu không phải tệp thử.

## Kết quả phần A (16/11/2026, bản cục bộ nạp vào khung bằng `cdp.mjs serve`)

### Thăm dò Word cho việc dựng văn bản hành chính (Word 16.0, WordApi 1.9, WordApiDesktop 1.5)

Bài: `kiem-thu/cdp-bai/tham-do-3-vbhc.js` và ba bài thăm dò nhỏ chạy trực tiếp trong khung. Đã đo, không làm trái:

| Điều thử | Kết quả |
|---|---|
| Ghi `section.pageSetup` (khổ A4, bốn lề, `differentFirstPageHeaderFooter`) | Được; không phải sửa đổi được theo dõi → công cụ tự ghi cách hoàn tác |
| Số trang: `insertField("Start", "Page")` trong đầu trang | Được (đầu trang đọc lại ra "1") |
| Bảng hai cột không viền (`getBorder("All").type = "None"`), đặt bề rộng cột | Được |
| Thêm đoạn thứ hai vào một ô bằng `insertParagraph(…, "After")` trên đoạn cuối của ô | **Rơi sang ô kế.** Phải chèn `"Before"` đoạn sẵn có của ô |
| Đường kẻ = viền dưới của một đoạn rỗng: `paragraph.borders.items[2]` (0 trên, 1 trái, 2 dưới, 3 phải; `getByLocation` không có) | Được; thu ngắn bằng thụt lề trái / phải |
| Viền dưới đặt trên đoạn duy nhất (hoặc đoạn cuối) của một ô | **Dính sang ô cùng hàng** (xóa viền ô trái làm mất đường kẻ ô phải) → đường kẻ không bao giờ là đoạn cuối ô |
| Chèn bảng ngay SAU một đoạn rỗng | **Bảng nuốt đoạn rỗng đó, thứ tự bị đảo**, định danh của đoạn rỗng trỏ sang ô đầu của bảng → không chèn bảng sau đoạn rỗng (đoạn neo rỗng thì chèn trước nó) |
| Từ chối phần chèn có bảng (Hoàn tác kiểu cũ) | Để lại một đoạn rỗng cạnh bảng → Hoàn tác của công cụ này gỡ thẳng bảng và đoạn theo định danh, theo thứ tự từ cuối lên |
| Định dạng (phông, cỡ, canh lề) khi đang theo dõi | Mỗi đoạn sinh một sửa đổi "Formatted" (51 sửa đổi cho một tờ trình, lề phải rất rối) → chèn có theo dõi, định dạng phần mới chèn khi đã tắt theo dõi: còn 3 sửa đổi |
| `firstLineIndent`, `spaceAfter`, `lineSpacing` (12 = một dòng) | Được |

### Mã

- `src/core/adminDoc.ts`: bản vẽ văn bản hành chính (hàm thuần): phần đầu, tên loại và trích yếu, Kính gửi, nội dung, phần ký, nơi nhận; danh sách chỗ còn trống; không bịa số, ngày, người ký.
- `src/office/wordAdmin.ts`: hai công cụ Word mới — `word_create_admin_document` (dựng trọn văn bản đúng thể thức trong một lần gọi; văn bản trống thì đặt A4, lề 20/20/30/20 mm, số trang ở đầu trang ẩn trang 1) và `word_set_page_format` (đặt khổ giấy, lề). `word_format_text` thêm thụt đầu dòng, cách đoạn, giãn dòng. Word nay 29 công cụ.
- Kỹ năng dựng sẵn thứ 18 `soan-thao-van-ban-hanh-chinh` (chỉ Word; tự nạp theo từ khóa; gọi bằng "/"); câu gợi ý trên màn chào; lời nhắc hệ thống và hướng dẫn nền của Word nhắc công cụ mới.
- `src/core/claims.ts` + `Chat.tsx`: cảnh báo "AI nói đã thay đổi tài liệu nhưng lượt này không có thao tác ghi nào", nút "Yêu cầu làm lại bằng công cụ" và "Đổi model". Cảnh báo chỉ là phần hiển thị (lưu cùng hội thoại), không vào lịch sử gửi model.
- `llm.ts`: nhận thêm câu lỗi "`tool calling` is not supported" của Groq là "model không gọi được công cụ".

### Bảng nghiệm thu phần A

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| A-a | Soạn công văn và tờ trình, model trả tiền và model miễn phí: đủ thành phần, đúng lề và cỡ chữ | **Đạt 3/4 trên mã cuối, 1/4 mới đạt trên bản dựng trước** — Chạy thật: công văn · Haiku 5.5 16/16 điểm đối chiếu (17 giây); tờ trình · Haiku 5.5 17/17; công văn · Nemotron 3 Ultra free 16/16 (63 giây). Tờ trình · Nemotron free: trên bản dựng trước đủ thành phần, đúng lề (15/17 theo bài đối chiếu cũ; hai điểm trượt là do bài đối chiếu và do model gõ lẫn một ký tự lạ làm lệch phông); chạy lại trên mã cuối thì **OpenRouter hết 50 lượt miễn phí trong ngày** → chưa chạy lại được. Cả bốn lần AI tự gọi `word_check_page_format`, kết quả "đúng chuẩn". Đối chiếu độc lập bằng mã đọc lại văn bản (cỡ chữ, kiểu chữ, canh lề, thụt đầu dòng, đường kẻ, lề, khổ giấy, số trang) và xem bản in PDF của tờ trình | `03a`, `03b`, `03c`; `03d` (bản dựng trước), `03e` (hết lượt) |
| A-b | Kỹ năng tự nạp theo yêu cầu, gọi được bằng "/", có gợi ý ở màn chào | **Đạt** — Kiểm thử máy (7 câu yêu cầu tự nạp, 2 câu không nạp nhầm, "/"); Chạy thật: cả bốn lần khung hiện "Kỹ năng áp dụng: soan-thao-van-ban-hanh-chinh" | `test/unit/dot3.test.ts`; `03a`–`03d` |
| A-c | Nói "đã sửa" mà không ghi: có cảnh báo; có ghi thật: không cảnh báo; có Kiểm thử máy | **Đạt** — Kiểm thử máy 7 nhóm ca (11 câu khẳng định, 14 câu dễ báo nhầm, kể lại sửa đổi của đối tác…). Chạy thật (Cerebras · gpt-oss-120b, ca dựng có chủ ý: bảo model chỉ trả lời "Đã sửa…" mà không gọi công cụ): khung hiện cảnh báo, văn bản không đổi; bấm "Yêu cầu làm lại bằng công cụ" → sửa thật 150% → 120% có theo dõi, lượt đó không cảnh báo; câu hỏi chỉ đọc không cảnh báo. Với Haiku 5.5 cùng yêu cầu: model trả lời "Chưa sửa gì…" → không cảnh báo (đúng). Bốn lần soạn văn bản (có ghi thật): không cảnh báo. **Chưa kiểm được:** ca tự nhiên với ERNIE 5.5 Flash free (hết lượt miễn phí hôm nay) | `04-canh-bao-noi-da-sua.json`, `04-canh-bao-khung.png` |
| A-d | Kiểm thử chặn số hiệu văn bản pháp luật và tên công cụ trong kỹ năng | **Đạt** — `npm test` 383/383 | Kiểm thử máy |
| A-e | Bảng Chạy thật so các model miễn phí; đổi mặc định sau khi anh chọn | **Chưa kiểm được đủ** — mới có: Nemotron 3 Ultra free gọi công cụ đúng 3/3 lần (Đợt 2) + 3 lần hôm nay; ERNIE 5.5 Flash free 1 lần nói "Đã sửa" mà không gọi công cụ (Đợt 2). Khảo sát các model miễn phí khác phải chờ lượt miễn phí có lại (7 giờ sáng) | `tasks/dot-2-bang-chung/09a`, `09b`; `03c` |
| — | Tự kiểm tra Word (thêm 2 bước: dựng công văn và tờ trình rồi hoàn tác; đặt lề rồi hoàn tác, định dạng đoạn) | **Đạt 36/36** trong phiên Word mới | `02-tu-kiem-tra-word-phan-A.json` |

### Soát mã độc lập và phát hành 4.2.1

- Anh chọn 16/11/2026 (ô tích chọn): **phát hành 4.2.1 ngay**; model miễn phí: **ưu tiên Nemotron, cảnh báo ERNIE** (bật "Miễn phí" ở nguồn chưa từng chọn model thì chọn Nemotron 3 Ultra free; ERNIE 5.5 Flash free có dấu ⚠ trong bảng chọn model; model OpenRouter anh đang nhớ giữ nguyên). Bằng chứng dấu ⚠: `05-dau-canh-bao-model.json`, `.png`.
- Soát mã độc lập: 2 cao, 8 trung bình, 6 thấp — đã sửa 15, để lại 1; báo cáo `soat-ma-doc-lap-phan-A.md`.
- Sau khi sửa, chạy lại trên bản cục bộ 4.2.1: `npm test` **395/395**, tsc sạch, đóng gói đạt; Tự kiểm tra Word **36/36** (`02`), Excel **22/22** (`06`), PowerPoint 13/14 rồi **14/14** (`07a`, `07b`, lỗi bố cục cũ); công văn · Haiku 16/16 (`03a`), tờ trình · Haiku 17/17 (`03b`), cảnh báo 3/3 ca (`04`). Các tệp `03c`, `03d` (Nemotron free) là của bản dựng trước lần sửa theo soát mã (hết lượt miễn phí nên chưa chạy lại).
- **Phát hành 4.2.1 xong (16/11/2026):** đẩy tiến thẳng `0ba6322..3e9bcb4` lên `main`; GitHub Actions đạt; `gh-pages` "Bản chạy thật v4.2.1 (main@3e9bcb4)"; Pages phục vụ `taskpane.0770303f.js` (trùng gói đóng cục bộ). Kiểm trên bản đã phát hành (khung tải từ GitHub Pages, đã tắt nạp bản cục bộ): đầu khung v4.2.1, Tự kiểm tra Word **36/36** (`10`), Chạy thật soạn công văn · Haiku 5.5 **16/16** (`11`). Đã dọn: xóa 21 hội thoại thử trong Lịch sử của Word (còn 4 hội thoại có từ trước), nguồn về Anthropic · Claude Opus 5.5, Groq về GPT OSS 20B, OpenRouter vẫn ERNIE như trước.
- **Giới hạn đã biết của 4.2.1:** công cụ đặt thể thức trang chỉ xử lý 30 phần (section) đầu, số trang chỉ đặt ở phần đầu; chữ ký số, dấu, phụ lục, bản sao chưa có trong công cụ dựng; nội dung do model viết vẫn phải soát (model miễn phí có lần tự ghi số hiệu luật); cảnh báo "nói đã sửa" là nhận diện câu chữ nên có thể sót câu lạ.

### Phát hiện phụ khi làm phần A

- **Manh mối cho vấn đề mở "Word từ chối chèn nhận xét"**: hôm nay trạng thái `AccessDenied (Range.insertComment)` (kèm `GeneralException` ở `insertContentControl`) xuất hiện hai lần, đều ngay sau khi chạy bài thăm dò có chèn bảng lúc đang theo dõi rồi xóa bảng / đoạn theo định danh với theo dõi tắt; trước đó trong cùng phiên Word, Tự kiểm tra có nhận xét vẫn đạt. Mở lại Word thì hết. Chưa tách được thao tác nào gây ra (việc B5).
- Groq (gói miễn phí): giới hạn 8.000 token mỗi phút nhỏ hơn một yêu cầu Word của add-in (khoảng 8.300 token với 29 công cụ) → mọi model Groq trả lỗi 413 trong Word. Câu lỗi còn nguyên tiếng Anh và có mã tổ chức của tài khoản; chưa sửa.
- Model miễn phí (Nemotron) có lần tự ghi số hiệu một đạo luật vào dòng căn cứ dù kỹ năng dặn không bịa: thể thức đúng nhưng nội dung vẫn phải soát.

## Bảng thăm dò B1 — kho chung và kênh lệnh

Chạy 16/11/2026 với Word, Excel, PowerPoint cùng mở (cùng tài khoản Office, bản cục bộ 4.2.1), bài `kiem-thu/cdp-bai/tham-do-3-kenh-lenh.js`; chỉ ghi / đọc khóa thử `aidn.thu.dot3.*` rồi xóa. Bằng chứng: `tasks/dot-3-bang-chung/08-tham-do-kenh-lenh.txt`.

| Điều thử | Kết quả |
|---|---|
| Ba ứng dụng dùng chung một kho `localStorage` | **Có.** Khóa ghi ở khung Word tới khung Excel và PowerPoint ngay; chiều ngược lại cũng vậy |
| Sự kiện `storage` tới khung ở ứng dụng khác | **Có**, trễ 0–2 ms, cả ba chiều đã thử (Word → Excel, Word → PowerPoint, PowerPoint → Word, PowerPoint → Excel) |
| `BroadcastChannel` giữa các khung | **Có**, trễ 0–1 ms |
| Khung ở ứng dụng nền (không có tiêu điểm) chạy `Excel.run` / `Word.run` / `PowerPoint.run` để GHI | **Được**: Excel ghi một ô ở sheet mới trong 32–82 ms, Word thêm một đoạn trong 31–36 ms, PowerPoint thêm một hộp chữ trong 4–8 ms |
| Cửa sổ đích bị thu nhỏ (Excel, Word thu nhỏ bằng COM, chờ 20 giây rồi gửi) | **Vẫn nhận và ghi được**, độ trễ như trên; `document.hidden` của khung vẫn là `false` |
| Hẹn giờ của khung nền có bị bóp không | **Không**: nhịp 1 giây đều (trung bình 1.000 ms, lớn nhất 1.016 ms trong khoảng 50 nhịp ở cả ba khung) |
| Trả lời về khung gửi | Qua khóa trả lời trong kho + sự kiện `storage`: 0–2 ms |

**Kết luận: kênh lệnh cục bộ dùng được, không cần phương án dự phòng.** Thiết kế chốt: "hộp thư" là các khóa trong `localStorage` (bền, khung nào mở sau vẫn đọc được trạng thái), báo tin bằng sự kiện `storage`, có thăm dò định kỳ làm lối dự phòng khi sự kiện bị lỡ; nhịp sống của từng khung cũng là một khóa. Không dùng `BroadcastChannel` làm kênh chính (không bền, không cho biết khung kia có mở không).

## Bảng thăm dò B3, B4 — PowerPoint

Chạy 16/11/2026 trong PowerPoint 16.0 (PowerPointApi 1.3–1.10 đều có), bài `kiem-thu/cdp-bai/tham-do-3-ppt.js`; bằng chứng `tasks/dot-3-bang-chung/09-tham-do-powerpoint.json`.

| Điều thử | Kết quả | Dùng cho |
|---|---|---|
| Ảnh slide: `slide.getImageAsBase64({ width })` | **Được**, PNG, 40 ms | `ppt_get_slide_image` |
| Màu chủ đề: `themeColorScheme.getThemeColor` (12 màu) | **Được** | `ppt_get_theme`, kiểm bài |
| Bố cục của mẫu và loại ô giữ chỗ (`layouts`, `placeholderFormat.type`) | **Được** | `ppt_get_theme`, `ppt_set_layout` |
| Định dạng một cụm chữ / một đoạn: `textRange.getSubstring`, `paragraphFormat`, `bulletFormat` | **Được**; đoạn ngắt bằng `\n` | `ppt_format_text` |
| Bảng có sẵn: `shape.getTable()`, đọc `values`, sửa ô, `rows.add` | **Được** | `ppt_edit_table`, tổng quan kèm nội dung bảng |
| Vùng chọn: `getSelectedShapes`, `getSelectedTextRangeOrNullObject`, `getSelectedSlides` | **Được** | `ppt_get_selection` |
| Đổi bố cục: `slide.applyLayout(layout)` | **Được** (ô tiêu đề được giữ) | `ppt_set_layout`, sửa lỗi bố cục của slide biểu đồ |
| Chèn ảnh: hình chữ nhật tô bằng ảnh (`fill.setImage(base64)`) | **Được về kỹ thuật**; **chưa làm thành công cụ** (cần đường đưa ảnh đính kèm trong khung chat tới công cụ — để đợt sau) | — |
| Phông chủ đề | **Không có API.** Đọc được gián tiếp: xuất một slide (`exportAsBase64`) rồi đọc tệp chủ đề trong gói (chỉ đọc) | `ppt_get_theme`, kiểm bài |
| Ghi chú diễn giả | **Không có API** (đối tượng Slide không có thuộc tính ghi chú) → **không làm**, AI được dặn nói rõ | — |
| Cập nhật số liệu biểu đồ đã có | **Không có API** (hình loại Chart không có phương thức nào về dữ liệu) → **không làm**; cách thay: tạo biểu đồ mới bằng `ppt_add_chart` rồi xóa biểu đồ cũ sau khi người dùng đồng ý | — |

## Phần B — mã và kết quả (bản cục bộ 4.3.0)

### Mã

- **Ghi chéo ứng dụng** (`src/core/crossapp.ts`, `src/office/crossExec.ts`, `write_other_app` trong `src/office/common.ts`): hộp thư lệnh trong kho dùng chung; **luôn hỏi ở khung gửi** trước khi lệnh rời đi (cờ `alwaysAsk`: hỏi cả khi anh tắt "hỏi trước thao tác rủi ro", và "Cho phép hết lần này" không áp dụng); hộp hỏi nêu ứng dụng đích, tên tệp và nội dung. Khung nhận chỉ chạy ba thao tác THÊM mới (Word: thêm đoạn ở cuối, ghi có theo dõi · Excel: tạo sheet mới, công thức từ ứng dụng khác bị ghi thành chữ · PowerPoint: thêm một slide), kiểm lại hình dạng / kích thước / hạn / đã chạy chưa, hiện mục "📥 Vừa ghi theo yêu cầu từ …" kèm nút Hoàn tác. Khung đích chưa mở, tắt nhận ghi, hoặc không trả lời: báo rõ, không treo (rút lệnh khi hết hạn).
- **Chỉ báo và công tắc**: đầu khung hiện "Excel ✓ · Word —" theo nhịp sống 3 giây; Cài đặt có công tắc kết nối theo từng ứng dụng và công tắc "Nhận ghi từ ứng dụng khác".
- **PowerPoint** (`src/core/deckcheck.ts`, `src/office/pptDesign.ts`): 7 công cụ mới — `ppt_get_slide_image` (ảnh slide gửi cho model ở một tin nhắn riêng + số đo bằng mã), `ppt_check_deck`, `ppt_get_theme`, `ppt_format_text`, `ppt_edit_table`, `ppt_get_selection`, `ppt_set_layout`. PowerPoint nay 21 công cụ. Slide biểu đồ tự chuyển về bố cục có sẵn của mẫu khi PowerPoint chèn bố cục lạ; slide có hình trùng tên thì công cụ bắt chỉ hình bằng số thứ tự.

### Bảng nghiệm thu phần B

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| B-1 | Từ PowerPoint: "lấy bảng tiến độ ở Excel đang mở, tạo 3 slide và 1 biểu đồ" chạy trọn, đúng số | **Đạt** — Chạy thật (Haiku 5.5, 33 giây): đọc bảng ở Excel, thêm 3 slide đúng mẫu của bài (dải màu, phông Georgia / Verdana): biểu đồ cột kế hoạch – thực tế 5 hạng mục (100/100, 80/62, 60/55, 35/20, 10/0), bảng chênh lệch (0, −18, −5, −15, −10), slide nguyên nhân. Đối chiếu bằng mã đọc lại bảng và xem ảnh slide. Lần chạy đầu lộ lỗi slide biểu đồ lệch mẫu → đã sửa rồi chạy lại trên bài mới | `20`, `20b-slide-11/12/13.png` |
| B-2 | Từ Excel: "ghi tóm tắt sang Word đang mở": đoạn mới xuất hiện, có hỏi xác nhận, hoàn tác được ở Word | **Đạt** — Chạy thật (Haiku 5.5, 11 giây): hộp hỏi "Ghi sang Word (tệp “aidn-thu-vbhc.docx”): thêm 1 đoạn…"; Word có thêm đúng 1 đoạn (1 sửa đổi được theo dõi), số liệu đúng (62/80, 20/35, 0/10, 55/60); khung Word hiện mục "Vừa ghi theo yêu cầu từ … Excel"; bấm Hoàn tác ở khung Word → văn bản về 2 đoạn ban đầu, 0 sửa đổi | `16` |
| B-3 | Đóng khung ở ứng dụng kia: AI báo rõ, không treo; chỉ báo đúng | **Đạt** — Chạy thật: đóng Word → chỉ báo ở Excel thành "Word — · PowerPoint ✓"; AI trả lời trong 10 giây "chưa ghi được sang Word… mở tệp Word, bấm nút AI Đa Nguồn…" và đưa đoạn tóm tắt để chép tay | `17`, `12` |
| B-4 | Slide cố ý để chữ tràn khung: AI tự phát hiện qua ảnh và sửa hết tràn | **Đạt** — Chạy thật (Haiku 5.5, 11 giây, 3 bước): chụp ảnh slide 6 → nới và dời khung "Rủi ro", giảm cỡ chữ 16 → 15 → chụp lại, hết tràn; ô tiêu đề và ô nội dung không bị đụng. Ảnh trước / sau đã xem. Lần chạy đầu AI sửa nhầm ô tiêu đề (hai hình trùng tên, phép đo "chồng nhau" báo cả khung rỗng) → đã sửa công cụ rồi chạy lại | `18`, `18a`, `18b` |
| B-5 | Bài 10 slide cài 5 lỗi lệch mẫu: kiểm bài tìm đủ 5 | **Đạt** — Tự kiểm tra (bước "Kiểm bài theo mẫu": cài 5 lỗi, tìm đủ 5, không báo thêm) trên bài trống và trên bài thử 10 slide; Kiểm thử máy cùng ca; Chạy thật trên `kiem-thu/ppt-thu.mjs`: tìm đủ 5 lỗi cài sẵn và cả slide chữ tràn | `15`, `21`, `19` |
| B-6 | Slide mới dùng đúng màu và bố cục của mẫu | **Đạt** — Tự kiểm tra (bước "Slide mới dùng đúng bố cục và màu của mẫu"): bố cục đọc lại trùng bố cục của mẫu, khối màu = Accent1 của mẫu (bài trống: #5B9BD5; bài 10 slide: #4472C4) | `15`, `21` |
| C-1 | `npm test` đạt hết, tsc sạch, đóng gói đạt | **Đạt** — `npm test` **420/420** (363 cũ + 57 mới), `npx tsc --noEmit -p .` sạch, `npm run build` đạt | Kiểm thử máy |
| C-2 | Mỗi công cụ mới có bước Tự kiểm tra; bước cũ không hỏng | **Đạt** — trên mã cuối (sau khi sửa theo soát mã): Word **37/37** (thêm bước ghi chéo), Excel **23/23**, PowerPoint **21/21** trên bài mới tinh (thêm 6 bước PowerPoint + bước ghi chéo). Bước "Biểu đồ gốc" nay kiểm điều người dùng thấy — slide biểu đồ dùng bố cục có sẵn của mẫu — và ghi rõ ở dòng kết quả rằng PowerPoint vẫn thêm 1 bố cục không dùng vào mẫu ở lần chèn đầu (add-in không có API xóa bố cục) | `13`, `14`, `15`, `21` |
| C-3 | Soát mã độc lập đã chạy, phát hiện đã sửa hoặc ghi lý do để lại | **Đạt** — phần A: 2 cao, 8 trung bình, 6 thấp, sửa 15, để lại 1; phần B: 0 cao, 6 trung bình + 1 khe cấu trúc, 10 thấp, sửa 11, để lại 6 | `soat-ma-doc-lap-phan-A.md`, `soat-ma-doc-lap-phan-B.md` |
| C-4 | Bằng chứng không lộ key, token, email, tên tài khoản Office, số dư | **Đạt** — soát bằng mã trước mỗi lần commit (mẫu key, email, mã tổ chức, tên trong hồ sơ Office): không có | lệnh soát trong nhật ký phiên |

### Phát hành 4.3.0 (16/11/2026, anh duyệt bằng ô tích chọn)

- Đẩy tiến thẳng `4659a6b..5717571` lên `main`; GitHub Actions đạt; `gh-pages` "Bản chạy thật v4.3.0 (main@5717571)"; Pages phục vụ `taskpane.6f0addd7.js` (trùng gói đóng cục bộ).
- Kiểm trên bản đã phát hành (ba khung tải từ GitHub Pages, đã tắt nạp bản cục bộ; tệp `23`): đầu khung v4.3, chỉ báo kết nối đúng ở cả ba khung; Tự kiểm tra **Word 37/37** (`24b`), **Excel 23/23** (`25`), **PowerPoint 21/21** (`26`); Chạy thật Excel ghi tóm tắt sang Word: có hỏi (hộp hỏi nêu tệp, số ký tự, nội dung), 1 sửa đổi được theo dõi, Hoàn tác ở Word về 0 (`27`).
- Dọn dẹp sau khi thử: xóa hội thoại thử ở cả ba khung (Word còn 4 hội thoại có từ trước, Excel 1, PowerPoint 2), nguồn về Anthropic · Claude Opus 5.5 (Groq: GPT OSS 20B, OpenRouter: ERNIE như trước), đóng Word, Excel, PowerPoint thử (không lưu), cổng 9229 đã đóng.
- Lần chạy Word đầu tiên trên bản phát hành ra 31/37 (`24a`) vì **cửa sổ Word đang thu nhỏ** — xem mục B5 ngay dưới: đây chính là nguyên nhân của vấn đề mở từ Đợt 2, không phải lỗi của bản 4.3.

### Sau Đợt 3 — bản 4.3.1: một nhiệm vụ, ít lượt gọi AI (anh nêu cuối phiên 16/11/2026)

Anh nêu: "giao AI một việc mà số lần gọi AI nhảy lên mấy chục lần; sợ phân mảnh lệnh làm AI không hiểu và làm sai mục tiêu — cần khắc phục."

- **Vì sao có nhiều lần gọi:** AI làm việc theo vòng "AI ra lệnh công cụ → add-in làm trên tài liệu → trả kết quả cho AI"; mỗi vòng là một lần gọi nhà cung cấp, và lần nào cũng gửi kèm toàn bộ yêu cầu của anh cùng mọi bước trước đó (AI không mất mạch, không phải nhiều lệnh rời nhau). Cái sai là **quá nhiều vòng**: bản 4.3.0 mỗi vòng chỉ ra một lệnh, đọc lại tổng quan sau từng bước và chụp từng slide nhiều lần.
- **Đã sửa (mã bản 4.3.1):** (1) lời nhắc chung của cả ba ứng dụng thêm mục "Một nhiệm vụ, ít lượt gọi nhất": gộp các lệnh không phụ thuộc nhau vào cùng một lượt, không đọc lại thứ đã có, xong thì trả lời ngay; (2) `ppt_get_slide_image` chụp được tới 4 slide trong một lần gọi; lời nhắc PowerPoint dặn làm xong hết rồi kiểm một lần; (3) khung hiện ngay dưới câu trả lời "1 nhiệm vụ · N lần gọi AI · Token …" kèm lời giải thích khi rê chuột.
- **Tiêu chí:** cùng nhiệm vụ "lấy bảng tiến độ ở Excel, tạo 3 slide và 1 biểu đồ" phải giảm ít nhất một nửa số lần gọi, kết quả vẫn đúng số, đúng mẫu; các việc khác không tăng; Tự kiểm tra không hỏng.

| Việc (Chạy thật, Claude Haiku 5.5) | 4.3.0 | 4.3.1 | Bằng chứng |
|---|---|---|---|
| PowerPoint: lấy bảng Excel → 3 slide + 1 biểu đồ | 17 lần gọi (270 nghìn token vào) và 30 lần gọi (513 nghìn) ở hai lần chạy | **8 lần gọi** ở cả hai lần chạy (161 và 156 nghìn token vào); số liệu đúng, slide đúng mẫu | `20`; `28`, `28b`, `28c` |
| PowerPoint: kiểm bài 10 slide | 5 lần gọi | **3 lần gọi** | `19`; số đo trong nhật ký phiên |
| Word: soạn một công văn đúng thể thức | 3 lần gọi | 3 lần gọi, 16/16 điểm đối chiếu | `32` |

- `npm test` **423/423**, tsc sạch, đóng gói đạt. Tự kiểm tra trên bản cục bộ 4.3.1: Word **37/37** (`31`), Excel **23/23** (`30`), PowerPoint **21/21** (`29`).
- Chưa làm ở bản này (để phiên sau, đã ghi HANDOFF): đo và giảm bước cho các việc Excel nhiều thao tác; trần số bước theo loại việc; công cụ gộp "thêm nhiều slide trong một lệnh"; đo với model miễn phí (model yếu có thể không gộp lệnh).

### Việc B5

- **Lỗi cũ "Biểu đồ gốc thêm một bố cục"**: đã chữa phần người dùng thấy — khi PowerPoint không ghép được bố cục và tự thêm bố cục lạ, slide biểu đồ được chuyển về bố cục có sẵn của mẫu (đúng dải màu, phông tiêu đề), ô giữ chỗ rỗng đè lên biểu đồ được gỡ. Bố cục thừa vẫn nằm lại trong mẫu, không dùng tới: PowerPoint không cho add-in xóa bố cục.
- **Phiên Word từ chối chèn nhận xét mới: ĐÃ TÌM RA NGUYÊN NHÂN (cuối phiên 16/11).** Khi **cửa sổ Word bị thu nhỏ** (WindowState = 2), Office.js từ chối `Range.insertComment` (`AccessDenied`) và `Range.insertContentControl` (`GeneralException`); khôi phục cửa sổ thì chạy ngay. Tái hiện cả hai chiều ba lần liên tiếp bằng cách đổi trạng thái cửa sổ qua COM (tệp `22b`); thêm nhận xét bằng COM của Word vẫn được, tài liệu không bị bảo vệ. Không phải lỗi của add-in. Hệ quả: (a) Tự kiểm tra Word phải chạy khi cửa sổ Word không thu nhỏ — `kiem-thu/mo-office.ps1` nay tự khôi phục cửa sổ; (b) người dùng thu nhỏ Word rồi nhờ AI (từ ứng dụng khác) thì thêm đoạn vẫn được, còn thêm nhận xét sẽ lỗi; **chưa làm**: câu báo lỗi tiếng Việt gợi ý "mở lại cửa sổ Word" khi gặp `AccessDenied` ở chèn nhận xét (việc nhỏ cho bản kế). Ghi chú lịch sử điều tra trước khi tìm ra: hôm nay trạng thái này xuất hiện hai lần sau các bài thăm dò; bài bisect `kiem-thu/cdp-bai/tham-do-3-nhan-xet.js` thử lần lượt 9 thao tác nghi ngờ (chèn / xóa bảng và đoạn lúc bật, tắt theo dõi; xóa theo định danh; `body.clear()`; bảng sau đoạn rỗng; từ chối mọi sửa đổi) trong một phiên Word mới — **không thao tác nào tái hiện** (tệp `22`). Cách hết vẫn là đóng mở lại Word.

### Giới hạn đã biết của 4.3.0

- Ghi chéo chỉ thêm mới (Word thêm đoạn ở cuối, Excel sheet mới, PowerPoint một slide ở cuối); không sửa nội dung có sẵn ở ứng dụng kia. Khung AI Đa Nguồn ở ứng dụng đích phải đang mở. Hai tài liệu cùng ứng dụng: đã có cơ chế chọn đúng tệp nhưng chưa thử trong Office với hai tài liệu thật.
- PowerPoint: không ghi chú diễn giả, không sửa số liệu biểu đồ đã có (không có API); chèn ảnh chưa thành công cụ; kiểm màu ngoài bảng màu chủ đề chưa bật; bài nhiều mẫu (master) chỉ xét mẫu đầu; chữ tràn khung là ước lượng (AI được dặn tin ảnh hơn số đo).
- Ảnh slide tốn token: một lượt tạo 3 slide có tự kiểm bằng mắt dùng khoảng 270 nghìn token đầu vào, phần lớn đọc từ bộ nhớ đệm (Anthropic). Model không nhận ảnh thì chỉ có số đo bằng mã.
- Chưa Chạy thật phần B với model miễn phí (hết lượt miễn phí của OpenRouter trong ngày).


## Bảng nghiệm thu

(điền khi có bằng chứng; bằng chứng ở `tasks/dot-3-bang-chung/`)
