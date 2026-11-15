# Đợt 0B — Độ bền, chi phí, an toàn (v3.6 → v3.7)

Lập 15/11/2026 cuối phiên Đợt 0A (bản nháp, rà lại trước khi làm). Tiêu chí nghiệm thu: bảng Đợt 0B trong `tasks/ke-hoach-v4.md` mục 5.

## Việc còn lại của Đợt 0A (làm trước)

1. Kiểm trong Office bản 3.6 (khung add-in Word đã mở được, quyền điều khiển Word/Excel/PowerPoint + msedgewebview2 + thanh tác vụ đã cấp 15/11):
   - Đầu khung hiện v3.6; ảnh bố cục mới, đo vùng hội thoại ≥ 70% chiều cao.
   - Enter gửi / Shift+Enter xuống dòng; gõ ba câu tiếng Việt có dấu (Unikey và Telex của Windows) không bị gửi giữa chừng.
   - Đổi qua lại hai nguồn: mỗi nguồn về đúng model đã dùng.
   - Kiểm tra nhà cung cấp chạy lại sau sửa "lỗi bị giấu chi tiết": OpenAI phải báo đúng nguyên nhân, không còn "Không kết nối được".
   - Sao lưu → xóa dữ liệu thử (chỉ dẫn thử, lệnh thử, hội thoại thử; KHÔNG xóa key thật) → khôi phục gộp. Thử tải tệp từ khung add-in ra đĩa; không được thì dùng "Chép vào bộ nhớ tạm".
   - Tự kiểm tra Word (17 bước: 14 cũ + 3 mới), Excel 22/22, PowerPoint 14/14.
2. Sửa `docs/so-sanh-claude-chinh-thuc.md` theo mục 7 của kế hoạch (chưa làm).
3. Chọn model mặc định cho OpenAI: `gpt-5.5` có thể cần xác minh tổ chức để stream → cân nhắc ưu tiên model phổ thông hơn hoặc thử không stream khi lỗi.

## Việc của Đợt 0B (theo kế hoạch mục 5)

Lịch sử chỉ thêm, không sửa · tự thử lại khi 429/5xx, không chạy lệnh công cụ bị cắt dở · bộ nhớ đệm lời nhắc Anthropic, hiện số token · ngưỡng tóm tắt theo sức chứa model · tách điều phối khỏi `Chat.tsx` (hook `useCatalog` đã tách phần danh mục) · an toàn (chỉ dẫn ẩn, cảnh báo rủi ro Word còn thiếu, key Gemini rời khỏi đường dẫn) · kiểm nốt 7 mục chưa chạy thật. Trước khi làm: rà nốt mã lõi tác tử (`loop.ts`, `llm.ts`, `compaction.ts`).
