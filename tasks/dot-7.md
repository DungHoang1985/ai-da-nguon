# Đợt 7 — bản 4.7.0 "trả nợ kỹ thuật" (kế hoạch, CHƯA làm)

Viết ngày 18/11/2026, gốc `origin/main` = `cb400ba`, bản đang chạy 4.6.0. Việc lấy từ mục "Việc kỹ thuật chưa làm" của `tasks/HANDOFF.md`. Đợt này không thêm tính năng lớn: làm nốt các việc đã hứa ở Đợt 3, 4 và vá chỗ bài thử từng gây hại.

Trạng thái: **chờ anh duyệt kế hoạch và trả lời 3 câu ở mục "Chờ anh chọn"**. Chưa sửa dòng mã nào.

Mọi bước kiểm trong Office chạy trên desktop ẩn (`kiem-thu/mo-office.ps1`). Anh đang mở Word / Excel / PowerPoint thì phiên dừng và hỏi, không mở bản thử.

## Hiện trạng đã đọc trong mã (18/11/2026)

- Trần số bước đã có nhưng cứng: `src/core/agent/loop.ts:98` đặt 20 bước cho mọi việc; chạm trần thì khung chỉ in một dòng chữ "Gửi "tiếp tục" nếu muốn AI làm tiếp" (`src/taskpane/components/Chat.tsx:840`), anh phải tự gõ.
- `ppt_add_slide` (`src/office/pptTools.ts:19`) mỗi lần gọi thêm đúng một slide → 3 slide là 3 lần gọi AI.
- Chèn ảnh vào slide: thăm dò Đợt 3 cho thấy `fill.setImage(base64)` chạy được; chưa có công cụ vì chưa có đường đưa ảnh đính kèm trong khung chat tới công cụ. Đợt 5 đã làm đường này cho PDF (`read_attachment_pages` giữ tệp trong bộ nhớ của phiên) — dùng lại.
- Lịch sử: `src/core/conversations.ts:19` giữ 60 hội thoại, hội thoại thứ 61 đẩy hội thoại cũ nhất ra và xóa hẳn (`:79–81`), không báo. Bài thử không có cách nào chạy mà không ghi lịch sử.
- Nhật ký Excel: thay hàng loạt, xóa trùng, tự điền, xóa vùng chỉ ghi một dòng nêu thao tác và vùng, không có giá trị cũ / mới; định dạng nhiều vùng chỉ ghi vùng đầu (mục 27 của bảng "Để lại" ở `tasks/dot-4.md`). Khóa ô công thức thay bố trí khóa sẵn có mà không hỏi (mục 9).
- Soát mã độc lập chưa chạy cho Đợt 0B, 5, 6.

## Việc

### A. Chạm trần số bước thì hỏi bằng nút, không bắt anh gõ

- Khi chạm trần: khung hiện một thẻ "Đã làm N bước, M thao tác ghi. Làm tiếp?" với nút **Làm tiếp** và **Dừng ở đây**. Bấm Làm tiếp = chạy tiếp cùng nhiệm vụ với một lượt trần mới; bộ đếm "1 nhiệm vụ · N lần gọi AI" cộng dồn, không tính là nhiệm vụ mới.
- Trần theo ứng dụng thay cho một số 20 chung (đề xuất, xem câu hỏi 1): Word 12, Excel 12, PowerPoint 20. Căn cứ là số đo Chạy thật đã có: soạn công văn 3 lần gọi, sửa một chỗ 3, việc Excel 6 yêu cầu 3–4, 3 slide + 1 biểu đồ 6–8.
- Không đoán "loại việc" từ câu chữ của anh: nhận diện câu chữ tiếng Việt đã báo nhầm nhiều ở Đợt 3 (cảnh báo "nói đã sửa"). Trần theo ứng dụng là thứ đo được.
- Hai lệnh liên tiếp bị công cụ từ chối với cùng lỗi: dừng sớm và báo, không đợi hết trần.
- Không đổi lời nhắc hệ thống (giữ bộ nhớ đệm).

### B. Thêm nhiều slide trong một lệnh

- Công cụ mới `ppt_add_slides`: nhận danh sách tối đa 10 slide (bố cục, tiêu đề, nội dung), thêm theo thứ tự trong một lần gọi; Hoàn tác gỡ cả lô như một thao tác. `ppt_add_slide` giữ nguyên.
- Lời dặn trong mô tả công cụ: từ 2 slide trở lên thì dùng công cụ này.

### C. Chèn ảnh đính kèm vào slide

- Công cụ mới `ppt_insert_image`: lấy một ảnh anh đã đính kèm trong hội thoại đang mở (theo tên hoặc số thứ tự), đặt vào slide bằng hình chữ nhật tô ảnh, giữ đúng tỷ lệ ảnh, vị trí và cỡ theo tham số hoặc tự vừa vùng nội dung. Có Hoàn tác.
- Giới hạn ghi rõ trong kết quả công cụ: chỉ ảnh đính kèm của phiên đang mở (tải lại khung là mất, phải đính lại); ảnh là phần tô của hình nên không cắt xén được như ảnh thường; không tải ảnh từ địa chỉ web.

### D. Nhật ký Excel và khóa ô

- D1. Thay hàng loạt, xóa trùng, tự điền, xóa vùng: nhật ký ghi giá trị cũ / mới từng ô, tối đa 200 ô mỗi thao tác; vượt thì ghi 200 ô đầu và một dòng "còn K ô nữa, không ghi chi tiết". Định dạng nhiều vùng: ghi đủ các vùng.
- D2. Khóa ô công thức trên sheet đã có bố trí khóa riêng: hỏi trước kèm lý do ("sheet đang có N ô mở khóa do người dùng đặt, thao tác này sẽ thay bố trí đó"). Sheet chưa có bố trí riêng thì chạy như cũ, không hỏi thêm.

### E. Bài thử không đụng lịch sử của anh

- E1. Chế độ thử: bài thử bật một cờ trong `sessionStorage` của khung (không có nút trên giao diện); khi cờ bật, hội thoại không được lưu và đầu khung hiện chữ "chế độ thử — không lưu lịch sử". Cờ mất khi đóng Office. `kiem-thu/cdp.mjs` thêm lệnh bật / tắt; các bài trong `kiem-thu/cdp-bai/` bật cờ ở đầu bài.
- E2. Cho người dùng thật (xem câu hỏi 2): khi lịch sử đã đủ 60 và sắp đẩy hội thoại cũ ra, khung báo một dòng "Lịch sử đã đầy, hội thoại cũ nhất “…” sẽ bị xóa" thay vì xóa lặng lẽ.

### F. Soát mã độc lập còn nợ

- Ba lượt, mỗi lượt một tác tử (không quá 4 tác tử cùng lúc), giao rõ quy tắc dự án, các điều đã đo trong Office và nhiệm vụ "cố tìm ca báo nhầm và ca lách được phép kiểm an toàn":
  - Đợt 0B: mã của `58555e1..0c570a7` đã bị viết lại nhiều ở các đợt sau → soát **trạng thái hiện tại** của `loop.ts`, `turn.ts`, `llm.ts`, `compaction.ts`, `errors.ts` (vòng lặp, thử lại, bộ nhớ đệm, lệnh bị cắt dở), không soát diff cũ.
  - Đợt 5: `git diff 57eb4ae..0c40ac7 -- src test` (đọc PDF theo trang, đường gửi ảnh trang, chữ gốc của đoạn chờ duyệt).
  - Đợt 6: `git diff 0c40ac7..01a6a13 -- src test` (phương ngữ, key dùng chung, nguồn tự khai — chú ý key đi tới đúng tên miền, không lộ qua câu lỗi).
- Phát hiện mức cao và trung bình: sửa trong đợt này. Mức thấp: sửa hoặc ghi lý do để lại.

## Không làm ở đợt này

- Các việc ở mục D của `tasks/dot-5.md` (bộ nhớ dài hạn, điền mẫu, connectors…) — chờ anh chọn riêng.
- Kiểm bù với Gemini, DeepSeek, MiniMax, Kimi, OpenRouter trả tiền, model miễn phí — cần anh nạp tiền hoặc có key; làm thành một buổi riêng.
- Rút gọn mô tả công cụ để giảm token mỗi lượt Word: đổi mô tả công cụ là đổi phần được nhớ đệm và có thể đổi cách model gọi công cụ, phải đo lại cả 5 bài hợp đồng — việc của một đợt riêng.
- Các mục "Để lại" còn lại của Đợt 3 và 4 (Hoàn tác tiêu đề trục biểu đồ, Pivot, vùng in nhiều vùng…): hiếm gặp, giữ nguyên lý do đã ghi.
- Ghi chéo với hai tài liệu cùng ứng dụng; thăm dò ba ứng dụng cùng lúc trên desktop ẩn: gộp vào buổi kiểm bù.

## Tiêu chí nghiệm thu

| # | Tiêu chí | Mức bằng chứng |
|---|---|---|
| A1 | Chạm trần: sự kiện mang đủ số bước và số thao tác ghi; "Làm tiếp" nối đúng lịch sử (chỉ thêm), bộ đếm cộng dồn; hai lệnh lỗi giống nhau liên tiếp thì dừng sớm | Kiểm thử máy |
| A2 | Trong Word, một việc cố ý dài hơn trần: thẻ hiện đúng số, bấm Làm tiếp thì AI làm nốt, bấm Dừng thì dừng và nút Hoàn tác vẫn đủ các thao tác đã làm | Chạy thật |
| A3 | Ba bài mốc không thụt lùi: soạn công văn ≤ 3 lần gọi AI, việc Excel 6 yêu cầu ≤ 4, không bài nào chạm trần | Chạy thật |
| B1 | `ppt_add_slides` thêm đúng thứ tự, đúng bố cục, Hoàn tác gỡ cả lô | Tự kiểm tra |
| B2 | Bài "lấy bảng Excel → 3 slide + 1 biểu đồ": ≤ 4 lần gọi AI (mốc 4.3.1: 6–8), đúng số, đúng mẫu khi xem ảnh slide | Chạy thật |
| C1 | `ppt_insert_image` đặt ảnh đúng tỷ lệ (sai lệch dưới 1%), trong khổ slide, Hoàn tác được | Tự kiểm tra |
| C2 | Đính một ảnh, nhờ "đặt ảnh này vào slide 2 bên phải": ảnh nằm đúng slide, không đè tiêu đề khi xem ảnh slide | Chạy thật |
| C3 | Không có ảnh đính kèm, hoặc đã tải lại khung: AI báo rõ cần đính lại, không bịa | Chạy thật |
| D1 | Thay hàng loạt 5 ô và xóa 3 dòng trùng: nhật ký có đủ giá trị cũ / mới từng ô; thao tác 300 ô: 200 dòng chi tiết + 1 dòng tổng; nội dung ghi vào nhật ký luôn là chữ (không chạy thành công thức) | Kiểm thử máy + Tự kiểm tra |
| D2 | Sheet có ô mở khóa riêng: có hộp hỏi kèm lý do, Bỏ qua thì bố trí khóa y nguyên; sheet thường: không hỏi | Tự kiểm tra |
| E1 | Cờ thử bật: chạy 3 hội thoại, số hội thoại trong lịch sử trước và sau bằng nhau; cờ tắt: lưu như cũ | Kiểm thử máy + kiểm bằng mã trong khung |
| E2 | (nếu anh chọn làm) Lịch sử đủ 60: có dòng báo trước khi hội thoại cũ nhất bị xóa | Kiểm thử máy + kiểm bằng mã trong khung |
| F1 | Ba báo cáo soát mã nằm trong `tasks/dot-7-bang-chung/`; mọi phát hiện cao và trung bình đã sửa hoặc có lý do để lại, sau đó chạy lại Tự kiểm tra | Báo cáo soát mã |
| G1 | Tự kiểm tra ba ứng dụng không thụt lùi (mốc 4.6.0: Word 38, Excel 39, PowerPoint 22; thêm các bước mới) | Tự kiểm tra |
| G2 | `npm test` đạt hết (mốc 524), kiểm kiểu sạch, đóng gói đạt | Kiểm thử máy |
| G3 | Thêm công cụ mới thì Chạy thật một bài với một model Gemini (bài học enum số); Gemini hết hạn mức thì ghi "Chưa kiểm được" | Chạy thật |

## Thứ tự làm

1. F trước (soát mã chạy nền bằng tác tử, phóng sớm để không chạm hạn mức cuối phiên), đồng thời E1 — vì mọi bài Chạy thật về sau cần cờ thử.
2. A, rồi B và C (cùng đụng mã PowerPoint), rồi D.
3. Sửa phát hiện của F, chạy lại Tự kiểm tra ba ứng dụng và các bài Chạy thật mốc.
4. Bảng kết quả vào cuối tệp này, cập nhật HANDOFF và bảng so sánh, xin anh duyệt phát hành 4.7.0. Sau phát hành: Tự kiểm tra lại trên bản phát hành.

Commit sau mỗi chặng trên nhánh riêng; chỉ đẩy `main` (tức phát hành) khi anh duyệt.

## Chờ anh chọn

1. **Trần số bước:** Word 12 · Excel 12 · PowerPoint 20 (đề xuất), hay giữ một số 20 chung và chỉ thêm nút Làm tiếp?
2. **E2 — báo trước khi lịch sử đầy:** làm (đề xuất), hay chỉ làm cờ cho bài thử?
3. **Model Chạy thật:** Claude Haiku 5.5 bằng key Anthropic như các đợt trước (tốn tín dụng API của anh; số lượt gọi của đợt này chưa ước được), hay anh muốn dùng nguồn khác?

## Việc buộc phải có anh

- Trả lời 3 câu trên và duyệt kế hoạch.
- Đóng Word / Excel / PowerPoint khi phiên cần mở Office thử (phiên sẽ hỏi trước).
- Duyệt phát hành 4.7.0.
