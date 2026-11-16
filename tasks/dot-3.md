# Đợt 3 — Soạn văn bản hành chính, cảnh báo "nói đã sửa", chéo ứng dụng và PowerPoint (4.2.0 → 4.3.0)

Viết ngày 16/11/2026, nhánh `claude/dot-3-b3b8f4` (gốc `origin/main` = `0ba6322`). Anh giao 16/11: "làm đợt 3 trong phiên mới. Xử lý luôn cả vấn đề thiếu kỹ năng soạn thảo văn bản hành chính".

Đợt gồm hai phần. **Phần A** làm trước (chữa đúng chỗ anh vừa gặp khi soạn Word), nhỏ và độc lập, có thể phát hành riêng thành 4.2.1 nếu anh chọn. **Phần B** là Đợt 3 của `tasks/ke-hoach-v4.md`.

## Trạng thái (cập nhật sau mỗi chặng)

- [ ] 0. Tệp này được commit trước khi viết mã
- [ ] A1 kỹ năng soạn văn bản hành chính · [ ] A2 cảnh báo "nói đã sửa mà không ghi" · [ ] A3 model miễn phí (đề xuất + hỏi anh)
- [ ] Kiểm phần A trong Word; hỏi anh phát hành 4.2.1 hay gộp 4.3.0
- [ ] B1 thăm dò kho chung và kênh lệnh (bảng kết quả ở dưới) · [ ] B1 mã ghi chéo · [ ] B2 chỉ báo và công tắc
- [ ] B3 PowerPoint phần chắc · [ ] B4 PowerPoint phần phải thử (bảng thăm dò ở dưới) · [ ] B5 hai việc tồn
- [ ] Soát mã độc lập, sửa, chạy lại · [ ] Bảng nghiệm thu, HANDOFF, CLAUDE.md, bảng so sánh · [ ] Hỏi anh phát hành, kiểm bản phát hành, dọn dẹp

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

## Bảng thăm dò B1 — kho chung và kênh lệnh

(chưa chạy)

## Bảng thăm dò B4 — PowerPoint

(chưa chạy)

## Bảng nghiệm thu

(điền khi có bằng chứng; bằng chứng ở `tasks/dot-3-bang-chung/`)
