# Đợt 5 — việc anh chọn ngày 17/11/2026

Anh chọn cả bốn việc, làm theo thứ tự: (A) đọc tài liệu quét theo trang · (B) kiểm nốt phần chưa kiểm · (C) hai việc nhỏ từ buổi so song song · (D) các việc khác của Đợt 5 (liệt kê để anh chọn, chưa làm).

Mọi bước kiểm trong Office chạy trên desktop ẩn (`kiem-thu/mo-office.ps1`, quy định 17/11/2026).

## A. Đọc tệp đính kèm theo trang

Nguồn kinh nghiệm: mục "Kinh nghiệm xử lý tài liệu quét" của `tasks/so-song-song.md`.

Thiết kế:
- Công cụ dùng chung `read_attachment_pages` (cả ba ứng dụng): đọc tệp PDF đính kèm theo số trang. Trang có lớp chữ (từ 300 ký tự) trả chữ; trang quét dựng ảnh (150 dpi, cạnh dài tối đa 1.600 điểm ảnh, JPEG) và gửi qua kênh ảnh công cụ sẵn có (`__images`, tin nhắn `fromTool`). Mỗi lần gọi tối đa 4 trang quét.
- Kết quả công cụ kèm lời dặn chép nguyên văn: đúng chữ in, không tóm lược khi được nhờ chép, chỗ không đọc được ghi `[?]`, không đoán; số và ngày chép đúng kiểu in; mọi số liệu lấy từ trang quét phải dẫn trang và nói rõ "đọc từ bản quét, cần đối chiếu bản gốc". Lời dặn nằm trong kết quả công cụ, không đổi lời nhắc hệ thống.
- Kết quả công cụ luôn ghi tổng số trang, trang đã đưa, trang còn lại (đếm trang vào bằng trang ra).
- Đường đi của PDF quét lúc gửi:
  - nguồn đọc được nguyên tệp và tệp từ 5 trang trở xuống: gửi nguyên tệp như cũ (đã Chạy thật 3/3 hôm 16/11);
  - còn lại (tệp dài hơn, hoặc nguồn không nhận PDF nhưng model nhận ảnh): không gửi tệp, AI đọc bằng công cụ;
  - không có công cụ (tắt "AI tự thao tác" hoặc model không gọi công cụ): như cũ.
- PDF lẫn trang quét: hiện báo "CHƯA được đọc, tách trang ra". Nay giữ tệp trong bộ nhớ của phiên và chỉ AI đọc các trang đó bằng công cụ.
- Cảnh báo riêng tư: khi gửi tệp đính kèm cho nguồn miễn phí hoặc qua OpenRouter, khung hiện một dòng nhắc (không chặn).

Không làm ở đợt này (ghi rõ): đọc hai lần độc lập rồi máy so số liệu; máy kiểm dãy Điều – khoản – điểm trên bản chép.

Tiêu chí nghiệm thu:

| # | Tiêu chí | Mức bằng chứng |
|---|---|---|
| A1 | Chọn trang, phân loại trang, đường đi của PDF quét, cảnh báo riêng tư đúng theo thiết kế | Kiểm thử máy |
| A2 | Công cụ dựng được ảnh trang quét và lấy được chữ trang có lớp chữ trong Office | Tự kiểm tra |
| A3 | PDF quét 3 trang, model nhận ảnh nhưng nguồn đi đường công cụ: đúng 3/3 số liệu, có dẫn trang | Chạy thật |
| A4 | PDF quét 3 trang, nguồn đọc nguyên tệp: vẫn đúng 3/3 như trước (không thụt lùi) | Chạy thật |
| A5 | Model không nhận ảnh: báo rõ không đọc được, không bịa | Chạy thật |

## B. Kiểm nốt phần chưa kiểm

| # | Việc | Mức bằng chứng |
|---|---|---|
| B1 | Gemini 3.8 Flash: 5 bài trên hợp đồng thử | Chạy thật |
| B2 | Gemini: PDF quét, Suy nghĩ kỹ | Chạy thật |
| B3 | Excel và PowerPoint trên desktop ẩn: khung mở, hẹn giờ, Tự kiểm tra | Tự kiểm tra |

## C. Hai việc nhỏ từ buổi so song song

| # | Việc | Tiêu chí |
|---|---|---|
| C1 | Đoạn đang có sửa đổi chờ duyệt | Khi AI đọc một đoạn có sửa đổi chờ duyệt, kết quả công cụ đưa cả chữ gốc và chữ sau sửa, kèm tên người sửa; bài "lãi chậm thanh toán" nêu được cả 45 và 90 ngày và nói rõ đang chờ duyệt |
| C2 | Lý do kèm mỗi chỗ sửa | Công cụ sửa Word nhận thêm lý do ngắn; dòng bước trên khung và mục Hoàn tác hiện lý do đó |

## D. Việc khác của Đợt 5 (chờ anh chọn)

Theo `tasks/ke-hoach-v4.md`: bộ nhớ dài hạn giữa các hội thoại · tải bản ghi hội thoại · gợi ý gạch chân bấm để chấp nhận trong Word · điền mẫu qua ô điền sẵn · connectors (MCP) · công cụ chạy mã Office.js do AI viết · phím tắt mở khung · đọc một trang web cụ thể · Outlook · Sign in with Poe · hàm `=AI(…)` trong ô Excel.

## Kết quả (17/11/2026, kiểm trên bản cục bộ; anh duyệt và đã phát hành bản 4.5.0 cùng ngày)

Kiểm thử máy 504/504, kiểm kiểu sạch, đóng gói đạt. Mọi bước trong Office chạy trên desktop ẩn. Bằng chứng: `tasks/dot-5-bang-chung/`.

| # | Kết quả | Bằng chứng |
|---|---|---|
| A1 | Đạt | Kiểm thử máy `test/unit/dot5.test.ts` (34 kiểm thử mới) |
| A2 | Đạt | Tự kiểm tra: Word 38/38, Excel 39/39, PowerPoint 22/22, đều có bước "Đọc tệp PDF đính kèm theo trang" (trang 1 ra 521 ký tự, trang 2 ra ảnh JPEG 13 KB) — `tu-kiem-tra-*.json` |
| A3 | Đạt | Chạy thật Sonnet 5.5, PDF quét 7 trang (3 trang thật xen 4 trang trắng) đi đường công cụ: đúng 3/3 số liệu, dẫn đúng trang 2, 4, 7, tự nêu các trang 1, 3, 5, 6 trống và câu "đọc từ bản quét, cần đối chiếu bản gốc"; 11 giây — `pdf-sonnet55-7trang.json` |
| A4 | Đạt | Chạy thật Sonnet 5.5, PDF quét 3 trang gửi nguyên tệp: đúng 3/3, 6 giây — `pdf-sonnet55-3trang.json` |
| A5 | Đạt sau một lần sửa | Lần 1 (`pdf-cerebras-7trang.json`): Cerebras `gpt-oss-120b` bị gửi ảnh trang và từ chối 400 ("Content type 'image_url' is not supported") vì danh mục không ghi model này không nhận ảnh. Đã sửa: nguồn không nhận PDF mà không ghi rõ là nhận ảnh thì không gửi ảnh trang. Lần 2 (`pdf-cerebras-7trang-lan2.json`): báo rõ không đọc được bản quét, gợi ý Anthropic / Gemini, 3 giây, không bịa số |
| B1 | Chưa kiểm được | Gemini 3.8 Flash: sáng 17/11 trả lời đúng một bài trên bản phát hành 4.4.1; chiều hết hạn mức (429) |
| B2 | Chưa kiểm được | `pdf-gemini-3trang.json`: 429 hết hạn mức; `pdf-gemini-7trang.json`: máy chủ im lặng, khung báo lỗi sau 150 giây (hạn chờ im lặng chạy đúng). Suy nghĩ kỹ: chưa chạy |
| B3 | Đạt | Excel và PowerPoint trên desktop ẩn: khung mở sau 1–2 giây, hẹn giờ 30 nhịp / 3 giây, 181 khung hình; Tự kiểm tra Excel 39/39, PowerPoint 22/22 (có bước chụp ảnh slide) |
| C1 | Đạt | Chạy thật Sonnet 5.5, bài "lãi chậm thanh toán": nêu Điều 6.4 có sửa đổi của "Đối tác" chưa duyệt, bản gốc 45 ngày, bản đang sửa 90 ngày — `c1-lan3-tom.json`. (Lần đầu chưa đạt vì AI chỉ dùng `word_find`; đã cho kết quả tìm kiếm kèm chữ gốc của đoạn.) |
| C2 | Đạt | Chạy thật Sonnet 5.5, bài thêm khoản sau Điều 8.2: dòng bước "Chèn điều khoản mới sau [¶44] — lý do: Bổ sung khoản mới về hồ sơ bàn giao theo yêu cầu." — `c2-nhan.json` |

Chưa kiểm: PDF lẫn trang quét với tệp thật (mới có kiểm thử máy cho phần ghép câu hỏi); dòng nhắc riêng tư trên khung (mới có kiểm thử máy); đường công cụ với một nguồn không nhận PDF nhưng được ghi rõ là nhận ảnh (chưa có nguồn nào như vậy đang dùng được); mục Hoàn tác kèm lý do (mã có, chưa nhìn trên khung).

Phát hành 4.5.0 (17/11/2026): `main@0c40ac7` → `gh-pages@35b174e`, gói `taskpane.b98efd63.js`; Tự kiểm tra trên bản phát hành ở desktop ẩn: Word 38/38, Excel 39/39, PowerPoint 22/22, đều có bước đọc PDF theo trang (`tasks/dot-5-bang-chung/phat-hanh-450-*.json`).

Kiểm nốt trên bản phát hành 4.5.0 (17/11/2026, Word trên desktop ẩn, Chạy thật):

| Việc | Kết quả | Bằng chứng |
|---|---|---|
| PDF lẫn trang quét với tệp thật (trang 1, 3 có chữ; trang 2 là ảnh quét) | Đạt | Sonnet 5.5 tự gọi `read_attachment_pages` cho trang 2; đúng 3/3 (48.750.000.000 đồng ở trang 2; 24 tháng ở trang 1; 8% ở trang 3), có câu "đọc từ bản quét… cần đối chiếu"; 7 giây, 2 lần gọi AI — `pdf-lan-trang-quet-sonnet55.json` |
| Dòng nhắc riêng tư trên khung | Đạt | Tệp ghi chú bịa gửi qua OpenRouter tới model miễn phí: khung hiện dòng "Tệp đính kèm được gửi qua OpenRouter tới một model miễn phí…", còn hiện sau khi AI trả lời — `nhac-rieng-tu-openrouter.json` |
| Hoàn tác sau chỗ sửa có lý do | Đạt một phần | Dòng bước và hộp hỏi hiện "Chèn nội dung (end) — lý do: …"; bấm "Hoàn tác thay đổi (1)" gỡ đúng đoạn vừa thêm, khung báo "Đã hoàn tác 1 thao tác". Nhưng nút Hoàn tác chỉ hiện số thao tác: nhãn kèm lý do của từng thao tác chỉ lộ ra trong câu báo lỗi khi hoàn tác hỏng — `c2-hoan-tac.json` |

Phát hiện phụ: PowerPoint thoát chậm hơn 20 giây sau lệnh đóng → `mo-office.ps1` nay chờ tới 45 giây.
