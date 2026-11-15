# Đợt 0B — Độ bền, chi phí, an toàn (v3.6 → v3.7)

Lập 15/11/2026 cuối phiên Đợt 0A (bản nháp, rà lại trước khi làm). Tiêu chí nghiệm thu: bảng Đợt 0B trong `tasks/ke-hoach-v4.md` mục 5.

## Việc còn lại của Đợt 0A — kết quả kiểm trong Office bản 3.6 (15/11/2026, trên máy)

Bằng chứng: `tasks/dot-0a-bang-chung/08`–`20`. Kiểm trên Word (Document1 trống), Excel (Book1), PowerPoint (Presentation1), Office 16.0.20527, khung add-in tải từ GitHub Pages.

| Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|
| Đầu khung hiện v3.6; bố cục một dòng đầu; vùng hội thoại ≥ 70% chiều cao | **Đạt** — khung 595 điểm ảnh cao, vùng hội thoại khoảng 470 (79%); không còn hàng điều khiển phía trên; nhãn model luôn hiện cạnh ô nhập | Ảnh 09 |
| Enter gửi, Shift+Enter xuống dòng | **Đạt** | Chạy thật trong Word (tin hai dòng gửi bằng Enter, AI trả lời) |
| Gõ ba câu tiếng Việt có dấu bằng Unikey (Telex) không bị gửi giữa chừng | **Đạt** — "Tôi đang thử gõ / Hợp đồng này dài / Điều khoản bảo mật" còn nguyên trong ô nhập, Enter mới gửi | Ảnh 10 |
| Như trên với bộ gõ Telex của Windows | **Chưa kiểm được** — máy chỉ cài bàn phím en-US, chưa có bộ gõ Vietnamese Telex của Windows; thêm bộ gõ là đổi cài đặt hệ thống nên không tự làm. Phần chặn khi đang ghép chữ (`isComposing`) mới có Kiểm thử máy | `Get-WinUserLanguageList` |
| Đổi qua lại hai nguồn: mỗi nguồn về đúng model đã dùng | **Đạt** — Anthropic ↔ Groq: Anthropic về Claude Haiku 5.5, Groq về GPT OSS 20B (không phải mục đầu danh sách) | Ảnh 11, 12 |
| Kiểm tra nhà cung cấp: OpenAI báo đúng nguyên nhân, không còn "Không kết nối được" | **Đạt** — "OpenAI trả về lỗi nhưng không cho add-in đọc chi tiết. Thường do: key sai hoặc hết hạn, tài khoản hết tiền, hoặc tài khoản chưa được dùng model gpt-5.5…". Từng nguồn: đạt 4/8 (Groq, Anthropic, Cerebras, Cohere); Gemini, DeepSeek hết tiền; OpenRouter trượt vì model đang nhớ | Tệp 08 |
| Sao lưu → xóa dữ liệu thử → khôi phục gộp | **Đạt** — tệp có 8 key, 1 chỉ dẫn riêng, 6 hội thoại, 1 lệnh; xóa chỉ dẫn thử, lệnh thử, hội thoại thử rồi "Gộp vào dữ liệu hiện có": cả ba trở lại, 8 key còn nguyên. Kỹ năng nhập thêm: chưa kiểm (máy chưa có kỹ năng nào nhập thêm) | Ảnh 13, 14, 15 |
| Tải tệp từ khung add-in ra đĩa | **Đạt** — tệp `AI-Da-Nguon-sao-luu-…-BI-MAT.json` nằm trong thư mục Tải xuống; không cần lối "Chép vào bộ nhớ tạm" | Ảnh 13 |
| Tự kiểm tra Word 17 bước | **Không đạt: 16/17** — bước "Tắt Track Changes rồi AI sửa: vẫn ghi có theo dõi" báo "2 → 2". Nguyên nhân: bước kiểm sửa ngay trong đoạn AI vừa chèn (đang là sửa đổi chờ duyệt) nên Word gộp, số sửa đổi không tăng. Sản phẩm đúng: chạy thật, Track Changes tắt, AI chèn rồi thay chữ đều ra sửa đổi có theo dõi | Ảnh 16, 17 |
| Tự kiểm tra Excel | **Đạt 22/22** | Ảnh 18 |
| Tự kiểm tra PowerPoint | **Đạt 14/14 ở lần chạy thứ hai**; lần đầu trên bài mới tinh 13/14: bước "Biểu đồ gốc" báo mẫu bị thêm bố cục (11 → 12) | Ảnh 19, 20 |
| Sửa `docs/so-sanh-claude-chinh-thuc.md` theo mục 7 | **Đạt** — 67 dòng, đã đếm lại cột trạng thái | `git diff` |

### Lỗi tìm ra khi kiểm, và việc đã làm

1. **Bước Tự kiểm tra Word đo sai** (không phải lỗi sản phẩm). Đã sửa `src/office/selftest.ts`: chèn một đoạn thường khi đang tắt theo dõi rồi mới nhờ công cụ sửa. **Chưa chạy lại trong Office** — phải phát hành rồi bấm Tự kiểm tra trong Word.
2. **`word_replace_text` không thay gì khi model truyền `max_replacements` là -1 hoặc số âm** mà vẫn báo "Đã thay -1 chỗ" (gặp với Groq · GPT OSS 20B). Đã sửa (`replaceLimit` trong `src/office/util.ts`), có 2 kiểm thử máy; `npm test` 120/120, kiểm kiểu sạch. **Chưa chạy lại trong Office.**
3. **PowerPoint: lần chèn biểu đồ đầu tiên trên bài mới tinh làm mẫu thêm một bố cục** dù mã đã gắn kiểu bố cục. Kết quả "14/14" của các bản trước nhiều khả năng đo trên bài đã có sẵn bố cục thừa. Chưa sửa → đưa vào Đợt 3 (hoặc sớm hơn nếu anh hay dùng biểu đồ).
4. **OpenRouter: model đang nhớ `thinkingmachines/inkling-2-small:free` bị OpenRouter chỉ cho dùng trong "agentic harnesses"**; add-in báo nhầm thành "Key không đúng hoặc không có quyền". Chưa sửa → gộp vào việc 3 dưới đây.
5. Nhỏ: bấm Gửi khi bảng Lịch sử đang mở thì bảng vẫn che hội thoại, phải bấm Đóng. Chưa sửa.

### Còn lại, chuyển sang Đợt 0B

1. Phát hành hai sửa đổi ở trên rồi chạy lại Tự kiểm tra Word (phải ra 17/17).
2. Kiểm Enter với bộ gõ Telex của Windows nếu anh có dùng (cần anh thêm bộ gõ trong Windows).
3. Chọn model mặc định: OpenAI `gpt-5.5` có thể cần xác minh tổ chức; OpenRouter tránh model bị giới hạn nơi dùng và báo đúng lý do khi gặp.
4. Tệp sao lưu thử còn trong thư mục Tải xuống (chứa 8 key thật và ba mục dữ liệu thử): giữ làm bản sao lưu hoặc xóa, tùy anh.

## Việc của Đợt 0B (theo kế hoạch mục 5)

Lịch sử chỉ thêm, không sửa · tự thử lại khi 429/5xx, không chạy lệnh công cụ bị cắt dở · bộ nhớ đệm lời nhắc Anthropic, hiện số token · ngưỡng tóm tắt theo sức chứa model · tách điều phối khỏi `Chat.tsx` (hook `useCatalog` đã tách phần danh mục) · an toàn (chỉ dẫn ẩn, cảnh báo rủi ro Word còn thiếu, key Gemini rời khỏi đường dẫn) · kiểm nốt 7 mục chưa chạy thật. Trước khi làm: rà nốt mã lõi tác tử (`loop.ts`, `llm.ts`, `compaction.ts`).
