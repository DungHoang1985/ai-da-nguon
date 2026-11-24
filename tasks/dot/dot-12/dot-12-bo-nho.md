# Đợt 12, phần E — thiết kế bộ nhớ dài hạn giữa các hội thoại

**Anh duyệt 23/11/2026**: cả ba câu ở cuối tệp chọn phương án khuyến nghị; anh dặn thêm lưu ý sức chứa của model miễn phí (Miễn phí: trần 800 ký tự, công cụ `remember` chỉ gửi khi câu có "từ giờ / ghi nhớ…"). Đã làm; kết quả và bảng nghiệm thu ở mục "Phần E" của `tasks/dot/dot-12/dot-12.md`.

## Vấn đề

Mỗi hội thoại mới, AI quên hết: cách anh muốn trình bày (kết luận trước, tiếng Việt, bảng Tiêu chí · Kết quả…), thuật ngữ quen dùng, tên dự án / các bên hay gặp, quy ước đánh số, mẫu câu trong công văn. Hiện có "Chỉ dẫn riêng theo ứng dụng" trong Cài đặt (`settings.instructions`), nhưng anh phải tự gõ và nó không tự lớn lên khi làm việc.

## Đã có sẵn để dựa vào

- `settings.instructions[host]` → đi vào lời nhắc hệ thống (`buildSystemPrompt`, mục chỉ dẫn riêng). Lời nhắc hệ thống phải giữ nguyên trong một hội thoại (bộ nhớ đệm lời nhắc — `test/unit/dot0b.test.ts`).
- Kho lưu cục bộ của add-in (`storage.ts`, nằm trong thư mục `Wef` của Office, tách theo tài khoản Office, không lên mạng).
- Cơ chế hỏi trước thao tác (hộp "AI muốn…") và chế độ thử không lưu lịch sử.

## Thiết kế đề xuất

1. **Nhớ gì:** các mục ngắn (mỗi mục ≤ 300 ký tự), loại: *cách làm việc* ("trả lời kết luận trước"), *thuật ngữ / tên riêng* ("Dự án H = Khu đô thị …"), *quy ước văn bản* ("công văn ký thừa lệnh Giám đốc"). **Không nhớ:** nội dung tài liệu, số liệu, thông tin mật, key — có bộ lọc chặn chuỗi giống key, số điện thoại, số CCCD, số tài khoản.
2. **Ai ghi:**
   - Anh ghi tay trong Cài đặt → mục "Bộ nhớ" (thêm, sửa, xóa, bật / tắt từng mục).
   - AI **đề xuất** ghi bằng công cụ `remember` (chỉ đề xuất): khung hiện thẻ "AI muốn nhớ: …  [Nhớ] [Bỏ]"; chỉ lưu khi anh bấm Nhớ. Câu "từ giờ luôn …" của anh là gợi ý để AI đề xuất.
3. **Dùng thế nào:** khi **mở hội thoại mới**, các mục đang bật (theo ứng dụng + chung) được ghép vào lời nhắc hệ thống, ngay sau chỉ dẫn riêng, tối đa khoảng 2.000 ký tự. Trong một hội thoại, bộ nhớ **đóng băng** (mục mới chỉ có hiệu lực từ hội thoại sau) → lời nhắc hệ thống không đổi giữa các lượt, bộ nhớ đệm giữ nguyên.
4. **Xem và xóa:** Cài đặt → Bộ nhớ: danh sách mục, ngày ghi, nguồn (anh / AI đề xuất), nút xóa từng mục và "Xóa hết". Sao lưu / khôi phục cùng bản sao lưu hiện có.
5. **Riêng tư:** lưu cục bộ; chỉ gửi đi như một phần lời nhắc tới nguồn anh đang dùng (giống chỉ dẫn riêng). Chế độ Miễn phí: thêm một dòng nhắc "bộ nhớ được gửi kèm tới nguồn miễn phí" khi bộ nhớ có mục.
6. **Token:** tối đa ~2.000 ký tự (≈ 550 token chữ tiếng Việt với Nemotron); chế độ Miễn phí tính vào trần lời nhắc.

## Tiêu chí nghiệm thu (khi làm)

- E1 Kiểm thử máy: bộ nhớ vào lời nhắc đúng chỗ, đóng băng trong hội thoại (kiểm thử chặn bộ nhớ đệm vẫn xanh), giới hạn độ dài, bộ lọc chặn key / số định danh (ca phải chặn và ca dễ chặn nhầm).
- E2 Tự kiểm tra (Word): thêm mục, hội thoại mới thấy trong lời nhắc, xóa mục.
- E3 Chạy thật: ghi "trả lời kết luận trước, dùng bảng Tiêu chí · Kết quả" → hội thoại mới, Nemotron trả lời theo đúng cách đó; AI đề xuất nhớ một quy ước, anh bấm Nhớ / Bỏ.

## Câu hỏi cần anh chọn

1. AI có được **đề xuất** nhớ (anh bấm Nhớ mới lưu) — Khuyến nghị — hay chỉ anh ghi tay?
2. Bộ nhớ **chung cho cả Word, Excel, PowerPoint** kèm mục riêng theo ứng dụng (Khuyến nghị), hay tách hẳn theo ứng dụng?
3. Mục mới có hiệu lực **từ hội thoại sau** (Khuyến nghị — giữ bộ nhớ đệm, rẻ hơn) hay ngay lượt sau (mất bộ nhớ đệm một lần)?
