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

## Kết quả

(điền khi làm xong từng phần)
