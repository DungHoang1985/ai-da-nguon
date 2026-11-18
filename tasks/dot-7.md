# Đợt 7 — bản 4.7.0 "trả nợ kỹ thuật" (đã phát hành trong bản 4.7.0, 18/11/2026)

> Trạng thái 18/11/2026: anh duyệt kế hoạch theo ba đề xuất mặc định. Mã của cả sáu việc đã viết và commit trên nhánh `claude/dot-7` (chưa đẩy `main`). Kiểm thử máy 560/560. Tự kiểm tra trên bản cục bộ: Word 38/38, Excel 40/40, PowerPoint 23/23. Chạy thật (Haiku 5.5) đạt các bài trừ B2 (3 slide + 1 biểu đồ còn 6 lần gọi AI, mốc đặt ra ≤ 4) và G3 (Gemini im lặng). Bảng kết quả và việc còn lại ở cuối tệp.

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

## Anh chọn (18/11/2026)

Duyệt kế hoạch theo ba đề xuất mặc định: trần Word 12 · Excel 12 · PowerPoint 20; làm E2 (báo khi lịch sử đầy); Chạy thật bằng Claude Haiku 5.5 qua key Anthropic.

## Đã làm (nhánh `claude/dot-7`, 4 commit từ `cb400ba`; chưa đẩy `main`)

- **A.** `src/core/agent/loop.ts`: `stepCapFor` (Word 12, Excel 12, PowerPoint 20); sự kiện `limit` kèm số lệnh ghi thành công; dừng sớm `stuck` khi một công cụ báo cùng một lỗi (y nguyên tham số: 2 bước; tham số có đổi: 3 bước). `Chat.tsx`: thẻ "AI đã làm N bước… Làm tiếp?" với nút Làm tiếp / Dừng ở đây; số lần gọi AI cộng dồn qua các lượt Làm tiếp. Ghi chú "add-in dừng lượt" nằm ở tin của add-in, không ở vai trợ lý (xem Bài học trong HANDOFF).
- **B.** `ppt_add_slides` (tối đa 10 slide một lệnh, một mục Hoàn tác cho cả lô); `ppt_add_slide` dùng chung hàm.
- **C.** `ppt_insert_image` + `src/core/imagefit.ts` (vị trí content / left / right / center / full hoặc left, top, width, height; giữ tỷ lệ; không ra ngoài slide; GIF, WebP và ảnh quá lớn được vẽ lại thành PNG). Ảnh đính kèm ở PowerPoint được giữ trong bộ nhớ của phiên và nêu tên trong tin nhắn.
- **D.** Nhật ký Excel: thay hàng loạt, xóa trùng, tự điền, xóa vùng ghi giá trị cũ / mới từng ô, tối đa 200 ô + một dòng tổng; định dạng nhiều vùng ghi đủ các vùng. `excel_protect` hỏi trước khi sheet đã có bố trí khóa riêng.
- **E.** Chế độ thử (`sessionStorage["aidn-che-do-thu"]`): không lưu lịch sử, đầu khung hiện "chế độ thử — không lưu lịch sử"; `cdp.mjs thu <cổng> on|off`, và `cdp.mjs eval` trên khung add-in tự bật (thêm `--luu` để không bật). Lịch sử đầy: dòng báo hội thoại cũ nhất bị xóa.
- **F.** Ba lượt soát mã độc lập (báo cáo trong `tasks/dot-7-bang-chung/soat-ma-*.md`) và các chỗ sửa — bảng dưới.
- Kiểm thử máy mới: `test/unit/dot7.test.ts` (19), `test/unit/dot7-soat.test.ts` (17). Tự kiểm tra thêm 1 bước PowerPoint, 1 bước Excel.

## Soát mã độc lập: đã sửa và để lại

| Lượt | Phát hiện | Đã sửa | Để lại |
|---|---|---|---|
| Lõi tác tử (trạng thái hiện tại) | 1 cao · 5 trung bình · 12 thấp | Cao: bản tóm tắt rỗng hoặc bị cắt cụt không còn thay hội thoại cũ. Trung bình: bản ghi đưa đi tóm tắt giữ phần cuối tin người dùng; Dừng hủy được lúc công cụ đang chạy; lệnh công cụ đứt luồng chưa có tham số không chạy với `{}`; trả lời rỗng hiện một dòng nói rõ; `model_context_window_exceeded` coi là bị cắt; cắt tóm tắt không rơi vào tin ảnh công cụ. Thấp: kết quả công cụ không đóng gói được; che đúng giá trị key đang dùng | Các mục thấp còn lại: 429 hết hạn mức theo ngày vẫn thử lại 3 lần; 408 và lỗi mạng tạm không thử lại; vài câu lỗi 400 / 401 / 403 báo sai loại; từ 10 lệnh song song thì mốc bộ nhớ đệm cũ ra ngoài tầm (chưa đo bằng API thật); `maskSecrets` che thừa ở vài câu thường |
| Đợt 5 | 3 cao · 6 trung bình · 9 thấp | Cao: trần chữ của `read_attachment_pages` hạ xuống 20.000 (dưới trần kết quả công cụ), trang quá dài bị cắt có ghi rõ, lời dặn đứng trước nội dung trang; PDF mà từ nửa số trang trở lên gần như không có chữ (bản quét mang dấu chữ ký số) được giữ tệp và ghi chú để AI đọc bằng ảnh; ảnh công cụ trong lịch sử có trần 16, cắt theo lô 8. Trung bình: nhắc riêng tư ở kênh Claude của OpenRouter; tham số `as_image`; kết quả tìm Word gắn chữ gốc trước khi tính ngân sách, chữ gốc trích quanh chỗ sửa; câu "không tìm thấy" nhắc có sửa đổi chờ duyệt; chỉ rút chữ 24 trang đầu của yêu cầu. Thấp: "1 - 3" được nhận | #4 phần còn lại: cờ "gói miễn phí" cho Groq, Cerebras, Gemini chưa có. #6 phần còn lại: chưa tìm thẳng trên chữ gốc (mới nhắc cách đọc bản gốc). Thấp #10, #12–#18 (trang chữ ngắn bị gọi là "scan" khi model không nhận ảnh; trần 15 MB cho bản quét đi đường công cụ; kho tệp đính kèm dùng chung giữa các hội thoại; lỗi pdf.js tiếng Anh…) |
| Đợt 6 | 0 cao · 2 trung bình · 9 thấp | Trung bình: nguồn có danh sách model cố định (Kimi) không còn hiện "Đã kết nối" khi key bị từ chối 401 / 403; key dạng JWT hoặc hex không tiền tố được che theo đúng giá trị đang dùng | 9 mục thấp: khối suy nghĩ cũ gửi lại cho nguồn nhóm 2 (có vòng gửi lại chế độ thường); câu lỗi "dán lại key" ở kênh dùng chung key; `max_tokens` của Claude 3.5 qua OpenRouter; nguồn tự khai luôn gửi `cache_control`; thẻ nguồn tự khai ẩn sau "xem thêm"; khôi phục gộp không gộp nguồn tự khai; cảnh báo số dư ở kênh Claude của OpenRouter; chuyển hướng mang `x-api-key`; nhận địa chỉ IP nội bộ |

Lượt soát Đợt 6 xác nhận: không có ca nào đưa key của nguồn này tới địa chỉ của nguồn khác; nguồn Anthropic gốc không đổi byte nào.

## Kết quả theo tiêu chí (18/11/2026, bản cục bộ, Office trên desktop ẩn)

| # | Kết quả | Bằng chứng |
|---|---|---|
| A1 | Đạt | Kiểm thử máy `dot7.test.ts` (6 kiểm thử nhóm A) |
| A2 | Đạt, làm trong **Excel** thay vì Word (anh mở Word giữa buổi) | Chạy thật Haiku 5.5, bài ghi 16 ô mỗi lệnh một ô: thẻ "AI đã làm 12 bước (hết một lượt), đã ghi 12 thao tác… Làm tiếp?", 12 ô đã ghi; bấm Làm tiếp → đủ 16 ô, "1 nhiệm vụ · 17 lần gọi AI" — `lam-tiep-excel-haiku55-lan3.json`. Hai lần chạy đầu (`…haiku55.json`, `…lan2.json`) lộ lỗi model nhận ghi chú dừng là lời mình; đã sửa. Chưa bấm thử nút "Dừng ở đây". Nút Hoàn tác vẫn theo từng lượt (12 rồi 4), không gộp |
| A3 | Đạt | Chạy thật Haiku 5.5: soạn công văn **2 lần gọi AI** (mốc ≤ 3), 16 điểm thể thức đạt — `cong-van-haiku55.json`; việc Excel 6 yêu cầu **4 lần gọi** (mốc ≤ 4), 10 lệnh, 0 lỗi, công thức, Pivot và biểu đồ đọc lại đúng — `excel-6-yeu-cau-haiku55.json`; không bài nào chạm trần |
| B1 | Đạt | Tự kiểm tra PowerPoint 23/23, bước "Đợt 7 · Thêm 3 slide…" — `tu-kiem-tra-powerpoint.json` |
| B2 | **Không đạt mốc** | Chạy thật Haiku 5.5, bài 3 slide + 1 biểu đồ (số liệu cho sẵn trong câu hỏi, không đọc từ Excel): nội dung đúng, AI dùng `ppt_add_slides` một lệnh cho các slide chữ, nhưng cả nhiệm vụ vẫn **6 lần gọi AI** (mốc đặt ra ≤ 4; mốc cũ 6–8). Ba lượt thừa là tự chụp ảnh kiểm, `ppt_check_deck` báo "2 phông lệch mẫu" ở tiêu đề slide 3 và 4, rồi sửa phông và dời biểu đồ — `ppt-3-slide-bieu-do-haiku55.json`. Chưa tìm vì sao tiêu đề slide mới bị báo lệch phông |
| C1 | Đạt | Cùng bước Tự kiểm tra: ảnh 400×200 đặt 430×215 ở nửa phải, đúng tỷ lệ, Hoàn tác gỡ được; ảnh không có thì báo rõ |
| C2 | Đạt | Chạy thật Haiku 5.5: đính `aidn-thu-so-do.png` (800×400), nhờ đặt vào slide 2 bên phải → hình "Ảnh aidn-thu-so-do.png" tại (520, 160) cỡ 400×200, đúng tỷ lệ 2:1, dưới tiêu đề (mép dưới tiêu đề 133); AI tự thu hẹp ô nội dung bên trái; 3 lần gọi — `ppt-dat-anh-haiku55.json`, `ppt-bai-sau-khi-thu.json` |
| C3 | Đạt | Chạy thật Haiku 5.5, chưa đính ảnh: AI gọi công cụ, nhận câu báo chưa có ảnh, trả lời "chưa thấy ảnh… vui lòng đính kèm lại", không bịa — `ppt-khong-co-anh-haiku55.json`. Ca "đã tải lại khung" chưa thử riêng (cùng đường mã) |
| D1 | Đạt | Kiểm thử máy (5 ô, 300 ô, công thức thành chữ) + Tự kiểm tra Excel 40/40, bước "Đợt 7 · Nhật ký…": thay 3 ô → 3 dòng W2, W4, W6 có cũ / mới; xóa trùng có giá trị cũ; Hoàn tác trả đủ — `tu-kiem-tra-excel.json` |
| D2 | Đạt | Cùng bước: sheet mặc định không hỏi; có ô mở khóa riêng thì hỏi; bố trí cũ còn nguyên sau khi chỉ hỏi |
| E1 | Đạt | Kiểm thử máy + trong khung: đầu khung hiện dòng chế độ thử; số hội thoại trong chỉ mục trước và sau ba lần chạy bài Chạy thật không đổi (0 → 0) |
| E2 | Một phần | Kiểm thử máy `nextToDrop` đạt; dòng báo trên khung chưa nhìn trong Office |
| F1 | Đạt | Ba báo cáo trong `tasks/dot-7-bang-chung/`; bảng "đã sửa và để lại" ở trên; 17 kiểm thử cho các chỗ sửa |
| G1 | Đạt | Bản cục bộ, desktop ẩn: **Word 38/38**, **Excel 40/40** (thêm 1 bước), **PowerPoint 23/23** (thêm 1 bước) — `tu-kiem-tra-word.json`, `tu-kiem-tra-excel.json`, `tu-kiem-tra-powerpoint.json`. Hai chỗ sửa theo soát mã đã Chạy thật trong Word: tìm "45 ngày" trên hợp đồng 54 trang có 30 sửa đổi chờ duyệt → AI nêu đúng Điều 6.4, bản gốc 45 ngày, đối tác đang sửa thành 90 ngày — `tim-45-ngay-haiku55.json`; đọc PDF quét 3 trang bằng `read_attachment_pages` với pages = "1 - 3" → đúng 3/3 số liệu, có dẫn trang và câu "cần đối chiếu bản gốc" — `pdf-theo-trang-haiku55.json` |
| G2 | Đạt | `npm test` 560/560, kiểm kiểu sạch, đóng gói đạt |
| G3 | Chưa kiểm được | Gemini 3.8 Flash trong PowerPoint: máy chủ im lặng, khung báo lỗi sau 150 giây (hạn chờ im lặng chạy đúng trên bản cục bộ) — `ppt-gemini-2-slide.json`. Model Gemini đang nhớ trên máy là "Gemini 2.5 Flash", Google báo không còn cấp cho người dùng mới; đã trả về đúng model đó như lúc đầu, anh nên chọn lại 3.8 Flash |

Dọn dẹp: sheet thử đã xóa, model trả về Anthropic · Claude Opus 5.5, Excel và PowerPoint thử đã đóng (Excel thử phải tắt bằng lệnh kết thúc tiến trình vì kịch bản đóng không chạy khi anh đang mở Word), chế độ thử tự mất khi Office đóng.

Sửa thêm sau các bài Chạy thật: nhãn bước của hai công cụ mới (trước hiện tên thô `ppt_add_slides`, `ppt_insert_image`).

Dọn dẹp sau buổi thứ hai: mọi bài chạy ở chế độ thử (không lưu lịch sử); model trả về Anthropic · Claude Opus 5.5 và Gemini · Gemini 2.5 Flash như lúc đầu; Word, PowerPoint, Excel thử đã đóng không lưu.

Phát hành 4.7.0 (18/11/2026, anh duyệt): `main@dea9bc7` → `gh-pages@8c61ba1`, lượt Actions 44233555510 xanh, GitHub Pages trả gói `taskpane.5d421d5a.js`. Tự kiểm tra trên bản phát hành ở desktop ẩn (đầu khung hiện v4.7): Word 38/38, Excel 40/40, PowerPoint 23/23 — `tasks/dot-7-bang-chung/phat-hanh-470-*.json`.

## Sau phát hành 4.7.0 — tìm nguyên nhân bài slide "phông lệch mẫu" (18/11/2026; đã sửa trên nhánh `claude/dot-7`, CHƯA phát hành)

Không phải `ppt_check_deck` báo nhầm. Hai lỗi thật, đều làm chữ trên slide mang Times New Roman thay cho phông của mẫu:

| Điều đo | Kết quả | Bằng chứng |
|---|---|---|
| Ghi chữ tiêu đề bằng Office.js, 9 kiểu chữ (`kiem-thu/cdp-bai/tham-do-7-phong-tieu-de.js`) | Chữ có ký tự **“ư”** ("Nhận xét xu hướng", "Xu huong tư") → PowerPoint đổi CẢ dòng sang Times New Roman. Chữ không dấu, có “ổ”, “ậ”, “ý” → giữ Calibri Light. Phông đọc ở mức cả khung vẫn báo Calibri Light; phải đọc theo từng dòng / ký tự mới thấy. Đặt lại tên phông sau khi ghi → về Calibri Light | `tham-do-phong-tieu-de.json`, `ppt-phong-theo-ky-tu.json` |
| Tiêu đề slide biểu đồ (`ppt_add_chart`, slide dựng bằng pptxgenjs rồi chèn vào) | Luôn Times New Roman, kể cả chữ không có “ư” | `ppt-phong-doc-bang-ma.json`, `ppt-phong-lech-mau-kiem.json` |

Đã sửa: `src/office/pptCore.ts` thêm `setTextKeepFont` và `restoreFont` (ghi chữ rồi trả phông cũ cho đúng những dòng bị đổi); dùng ở `ppt_add_slide`, `ppt_add_slides` (tiêu đề, nội dung), `ppt_set_shape_text`; `ppt_add_chart` trả phông ô tiêu đề của bố cục cho tiêu đề slide biểu đồ. Tự kiểm tra PowerPoint thêm bước "Đợt 7 · Chữ có “ư”…".

| Kiểm | Kết quả | Bằng chứng |
|---|---|---|
| Kiểm thử máy, kiểm kiểu, đóng gói | 560/560, sạch, đạt | — |
| Tự kiểm tra PowerPoint trên bản cục bộ đã sửa | **24/24**; bước mới: tiêu đề và nội dung có “ư” ra Calibri Light / Calibri, sau `ppt_set_shape_text` vẫn vậy, tiêu đề slide biểu đồ Calibri Light | `tu-kiem-tra-powerpoint-sau-sua-phong.json` |
| Chạy thật bài 3 slide + 1 biểu đồ (Haiku 5.5), hai lần | **5 lần gọi** rồi **4 lần gọi** (trước khi sửa: 6); mọi tiêu đề và nội dung đúng phông mẫu, AI không còn sửa phông | `ppt-3-slide-bieu-do-sau-sua-phong-lan1.json`, `…lan2.json`, `ppt-phong-sau-sua-lan1.json`, `…lan2.json` |

Chưa kiểm: Tự kiểm tra Word và Excel trên bản đã sửa (mã đổi chỉ nằm trong các tệp PowerPoint). Chưa sửa: lượt thừa còn lại ở bài này là AI đặt lại vị trí ô tiêu đề của slide biểu đồ, vì tổng quan báo ô đó ở toạ độ 0, 0, 0, 0 (ô giữ chỗ thừa hưởng vị trí từ bố cục). Các chỗ ghi chữ khác của PowerPoint (`ppt_add_textbox`, `ppt_add_shape`, `ppt_edit_table`, sơ đồ quy trình) chưa đo với chữ có “ư”.
## Việc còn lại

1. Cập nhật bảng so sánh `docs/so-sanh-claude-chinh-thuc.md` theo bản 4.7.0 (chưa làm).
2. Phát hành bản vá phông (4.7.1) khi anh duyệt. Sau đó: ô tiêu đề slide biểu đồ báo toạ độ 0, 0, 0, 0 (còn làm AI tốn một lượt); đo chữ có “ư” ở ô chữ, hình, bảng, sơ đồ.
3. E2: nhìn dòng báo "lịch sử đầy" trên khung (mới có kiểm thử máy). A2: bấm thử nút "Dừng ở đây".
4. G3: chạy lại một bài công cụ mới với Gemini khi máy chủ Gemini trả lời.
5. Các mục soát mã để lại (bảng ở trên).
