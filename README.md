# AI Đa Nguồn – Add-in Office (Word, Excel, PowerPoint)

Trợ lý AI trong khung bên phải của Office, dùng API key của nhiều nhà cung cấp:
OpenRouter, Google Gemini, Groq, Anthropic (Claude), OpenAI, Mistral, Cerebras, DeepSeek, SambaNova, NVIDIA NIM, Cohere.

- Nhập key trong thẻ **Cài đặt** → key chỉ lưu trên máy (bộ nhớ cục bộ của add-in).
- Tự liệt kê model theo nhà cung cấp, tự làm mới sau 24 giờ; lọc "Miễn phí" và "Chỉ model thao tác được" (OpenRouter, biểu tượng 🛠).
- **AI tự thao tác trong tài liệu** (chế độ tác tử, gọi công cụ):
  - Excel: đọc sổ tính, ghi giá trị/công thức, định dạng, tạo sheet, bảng (Table), biểu đồ, sắp xếp, xóa vùng — có nút **Hoàn tác**.
  - Word: thay chữ, viết lại đoạn (chỉ sửa phần khác), chèn tiêu đề/đoạn/gạch đầu dòng, định dạng, nhận xét, bảng — ghi bằng **Theo dõi thay đổi**.
  - PowerPoint: tạo slide theo bố cục có sẵn, sửa chữ, thêm ô chữ/hình/bảng — có nút **Hoàn tác**.
  - Từng bước hiện trong khung chat; nút Dừng; tùy chọn hỏi trước khi sửa.
- **Tự kiểm tra** (thẻ Cài đặt): chạy thật mọi thao tác trên tài liệu đang mở rồi dọn sạch, không cần key.
- Chế độ chỉ trò chuyện + nút **Chèn vào** khi tắt "AI tự thao tác" hoặc model không gọi được công cụ.

Bản chạy thật: https://dunghoang1985.github.io/ai-da-nguon/ (manifest: `dist/manifest.xml` sau khi build).

## Cài vào Office trên Windows (một lần)

Bấm đúp `cai-dat/cai-dat.cmd`: tải manifest bản chạy thật về `%LOCALAPPDATA%\AiDaNguon` và đăng ký với Office (khóa `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, giá trị = Id trong manifest).
Đóng hẳn rồi mở lại Word/Excel/PowerPoint → thẻ **Home** → **Mở AI Đa Nguồn**. Gỡ: `cai-dat/go-cai-dat.cmd`.

## Cập nhật bản chạy thật

`npm ci && npm run build`, rồi đưa nội dung `dist/` (bỏ `*.map`) lên nhánh `gh-pages`. Office tự dùng bản mới khi mở lại khung add-in.

## Phát triển

`npm ci` · `npm test` · `npm run build` · `npm start` (chạy thử bản localhost).
