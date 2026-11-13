# AI Đa Nguồn – Add-in Office (Word, Excel, PowerPoint)

Trợ lý AI trong khung bên phải của Office, dùng API key của nhiều nhà cung cấp:
OpenRouter, Google Gemini, Groq, Anthropic (Claude), OpenAI, Mistral, Cerebras, DeepSeek, SambaNova, NVIDIA NIM, Cohere.

- Nhập key trong thẻ **Cài đặt** → key chỉ lưu trên máy (bộ nhớ cục bộ của add-in).
- Tự liệt kê model theo nhà cung cấp, tự làm mới sau 24 giờ; lọc "Chỉ miễn phí" (OpenRouter).
- Chat trả lời chạy dần, có nút Dừng; gửi kèm nội dung đang chọn.
- **Chèn vào**: Excel → ô đang chọn (bảng Markdown được tách thành ô) · Word → sau chỗ con trỏ · PowerPoint → ô chữ/slide đang chọn.

Bản chạy thật: https://dunghoang1985.github.io/ai-da-nguon/ (manifest: `dist/manifest.xml` sau khi build).

## Cài vào Office trên Windows (một lần)

Tải `manifest.xml` của bản chạy thật về một thư mục cố định, rồi ghi một khóa registry trỏ tới tệp đó:

```
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AiDaNguon /t REG_SZ /d "D:\ADD IN\manifest.prod.xml" /f
```

Đóng hẳn rồi mở lại Word/Excel/PowerPoint → thẻ **Home** → **Mở AI Đa Nguồn**.
Gỡ: `reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AiDaNguon /f`.

## Phát triển

`npm ci` · `npm test` · `npm run build` · `npm start` (chạy thử bản localhost).
