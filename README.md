# AI Đa Nguồn – Add-in Office (Word, Excel, PowerPoint)

Trợ lý AI trong khung bên phải của Office, làm việc như add-in Claude cho Office nhưng dùng API key của nhiều nhà cung cấp:
OpenRouter, Google Gemini, Groq, Anthropic (Claude), OpenAI, Mistral, Cerebras, DeepSeek, SambaNova, NVIDIA NIM, Cohere.

- Nhập key trong thẻ **Cài đặt** → key chỉ lưu trên máy (bộ nhớ cục bộ của add-in).
- Tự liệt kê model theo nhà cung cấp, tự làm mới sau 24 giờ; lọc "Miễn phí" và "Chỉ model thao tác được" (OpenRouter, biểu tượng 🛠).

## AI tự thao tác trong tài liệu (chế độ tác tử)

- **Excel:** đọc sổ tính, ghi giá trị/công thức, định dạng, sheet, Table, biểu đồ, sắp xếp, xóa vùng; tìm kiếm, quét lỗi công thức,
  lần ô tham chiếu/phụ thuộc, lọc, PivotTable, định dạng có điều kiện, danh sách chọn, ghi chú ô, chèn/xóa hàng cột,
  đổi tên sheet, cố định dòng/cột — có nút **Hoàn tác**.
- **Word:** thay chữ, viết lại đoạn (chỉ sửa phần khác), chèn tiêu đề/đoạn/gạch đầu dòng, định dạng, bảng;
  đọc/trả lời/đánh dấu đã xử lý **nhận xét**; đọc và chấp nhận/từ chối **sửa đổi được theo dõi** (bản sửa của đối tác);
  đọc và điền bảng có sẵn; chuyển tới đoạn — mọi sửa đổi ghi bằng **Theo dõi thay đổi**.
- **PowerPoint:** tạo slide theo bố cục có sẵn, sửa chữ, ô chữ/hình/bảng; **biểu đồ gốc** sửa được dữ liệu (theo màu chủ đề);
  **sơ đồ quy trình** tự dàn trang; nhân bản/di chuyển/xóa slide; sửa vị trí, màu, chữ của hình — có nút **Hoàn tác**.
- Luôn hỏi trước khi ghi đè ô có dữ liệu hoặc xóa (ô, đoạn, slide, sheet); tùy chọn hỏi trước mọi thao tác sửa.
- Địa chỉ ô (`Sheet1!C12`), đoạn (`[¶12]`), slide (`slide 3`) trong câu trả lời bấm được để chuyển tới.

## Kỹ năng (Skills), lệnh "/" và các tính năng chung

- **17 kỹ năng tích hợp sẵn**, chuyển thể (viết lại tiếng Việt, không chép nguyên văn) từ skill chính thức của Anthropic
  `xlsx`, `docx`, `pptx`, `dataviz` và plugin `legal` (review-contract, triage-nda, legal-risk-assessment, compliance-check,
  signature-request, meeting-briefing, legal-response; Apache-2.0): rà soát hợp đồng, sàng lọc NDA, đánh giá rủi ro pháp lý,
  xử lý nhận xét, tóm tắt sửa đổi của đối tác, kiểm tra tuân thủ, kiểm tra trước khi ký, chuẩn bị họp, soạn phản hồi pháp lý,
  chuẩn hóa/kiểm tra lỗi bảng tính, mô hình tài chính dự án, phân tích dữ liệu, thiết kế biểu đồ, báo cáo tiến độ dự án,
  trình chiếu từ tài liệu, kiểm tra bài trình chiếu. AI **tự chọn** kỹ năng theo câu hỏi; gõ `/` để gọi trực tiếp.
- Nhập thêm kỹ năng theo chuẩn `SKILL.md` (dán, tệp `.md`, `.zip`/`.skill`, cả gói plugin Claude có `commands/`); "Lệnh của tôi" với `$ARGUMENTS`.
- **Chỉ dẫn riêng** cho từng ứng dụng (luôn áp dụng).
- **Lịch sử hội thoại** lưu trên máy (mở lại, xóa); tự tóm tắt khi hội thoại quá dài.
- **Tệp đính kèm:** PDF, DOCX, TXT/CSV/MD/JSON, ảnh.
- **Đọc chéo** nội dung đang mở giữa Excel – Word – PowerPoint.
- **Tìm trên web** (OpenRouter, Anthropic, Gemini).
- **Tự kiểm tra** (thẻ Cài đặt): chạy thật mọi thao tác trên tài liệu đang mở rồi dọn sạch, không cần key.

Bản chạy thật: https://dunghoang1985.github.io/ai-da-nguon/ (manifest: `dist/manifest.xml` sau khi build).

## Cài vào Office trên Windows (một lần)

Bấm đúp `cai-dat/cai-dat.cmd`: tải manifest bản chạy thật về `%LOCALAPPDATA%\AiDaNguon` và đăng ký với Office (khóa `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, giá trị = Id trong manifest).
Đóng hẳn rồi mở lại Word/Excel/PowerPoint → thẻ **Home** → **Mở AI Đa Nguồn**. Gỡ: `cai-dat/go-cai-dat.cmd`.

## Cập nhật bản chạy thật

`npm ci && npm run build`, rồi đưa nội dung `dist/` (bỏ `*.map`) lên nhánh `gh-pages`. Office tự dùng bản mới khi mở lại khung add-in (trang HTML được lưu tạm tối đa 10 phút).

## Phát triển

`npm ci` · `npm test` · `npm run build` · `npm start` (chạy thử bản localhost).
