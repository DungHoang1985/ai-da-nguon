# 04 — Chạy thật: key không có số dư (Gemini), 16/11/2026

Mức bằng chứng: **Chạy thật** trong Word (tệp thử `aidn-thu-dot2.docx`, bản cục bộ 4.2.0 nạp vào khung, đầu khung v4.2). Bài: `kiem-thu/cdp-bai/dot-2-het-tien.js`. Tài khoản Gemini của anh đang hết tiền trả trước; không cần anh làm gì. Bài không đọc kho lưu; bộ đếm chỉ ghi tên miền, đường dẫn và mã trạng thái.

## Trình tự và kết quả

1. **Anthropic · Claude Haiku 5.5** — "Không dùng công cụ. Ghi nhớ mã thử của phiên này: SEN-4207. Trả lời đúng hai chữ: đã nhớ." → "đã nhớ" (Token: vào 14.197 · ra 91 · ghi vào bộ nhớ đệm 14.193).
2. Đổi sang **Google Gemini · Gemini 2.5 Flash**, gửi "Không dùng công cụ. Câu thử hết tiền: một tuần có mấy ngày?".
   - Các lần gọi tới Gemini: `generativelanguage.googleapis.com/v1beta/models/<model>:streamGenerateContent -> 402` — **đúng 1 lần**, cả lượt mất 2 giây (không tự thử lại).
   - Thông báo trong khung chat:

     > Google Gemini: tài khoản đã hết tiền (hết tín dụng) nên yêu cầu bị từ chối. Hội thoại vẫn còn nguyên: đổi sang nguồn khác ở nhãn model cạnh ô nhập rồi gửi lại, hoặc nạp thêm ở trang của Google Gemini. (Your prepayment credits are depleted. Please go to AI Studio at https://ai.studio/projects to manage your project and billing. Learn more at https://ai.google.dev/gemini-api/docs/billing#prepay. )

   - Nút cạnh thông báo: `Đổi nguồn` · `Mở Cài đặt` · `Chép chi tiết lỗi`.
   - Câu hỏi của lượt 1 và lượt 2 vẫn hiện trong khung.
3. Bấm `Đổi nguồn` → bảng chọn nhà cung cấp và model mở ra. Chọn lại Anthropic · Claude Haiku 5.5.
4. "Không dùng công cụ. Mã thử tôi nhờ ghi nhớ ở đầu hội thoại này là gì? Chỉ ghi đúng mã đó." → **"SEN-4207"** (Token: vào 15.564 · ra 63 · đọc từ bộ nhớ đệm 14.193 · ghi vào bộ nhớ đệm 1.367).

Ảnh khung sau bước 4: `04b-khung-chat-sau-loi-het-tien.png`.

## Đối chiếu tiêu chí 5

| Yêu cầu | Kết quả |
|---|---|
| Thông báo tiếng Việt, rõ lý do | Có: "tài khoản đã hết tiền (hết tín dụng)", nêu tên nguồn, nêu việc nên làm |
| Hội thoại còn nguyên | Có: ba lượt đều còn trong khung; sau lượt lỗi, đổi nguồn rồi hỏi lại thì model trả đúng mã của lượt 1; lượt 4 đọc lại 14.193 token từ bộ nhớ đệm (phần đầu lịch sử không bị đổi) |
| Gợi ý đổi nguồn | Có: trong câu thông báo và nút `Đổi nguồn` mở bảng chọn |

Chạy lại sau khi sửa theo lần soát mã độc lập (commit `e352118`): kết quả như trên — 1 lần gọi, 402, 2 giây, cùng câu thông báo, nút Đổi nguồn mở bảng chọn, model trả đúng "SEN-4207" (đọc 14.193 token từ bộ nhớ đệm). Tệp `08e-chay-lai-het-tien-gemini.json`.

Ghi nhận: Gemini nay trả mã **402** cho ca này (kế hoạch và phiên trước ghi 429). Mã nào thì add-in cũng nhận ra qua câu "prepayment credits are depleted" và không thử lại (có kiểm thử máy cho cả hai mã).
