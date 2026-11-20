# Soát mã độc lập Đợt 11 (20/11/2026)

Người soát: tác tử soát mã riêng (model Sonnet), chỉ đọc; chạy `npx tsc --noEmit -p .` (sạch) và `npm test` (697/697 lúc soát). Đã quét bằng mã: lời nhắc Miễn phí Excel, PowerPoint không nêu tên công cụ ẩn; chuẩn làm việc không nêu tên công cụ nhóm; mô tả công cụ lõi không nhắc công cụ ẩn; kỹ năng nhắc công cụ chuyên đều kéo đúng nhóm; cùng id nhóm thì cùng nhãn; không công cụ nào thuộc hai nhóm; bộ công cụ không đổi trong một bước; enum là chuỗi.

Mọi mục dưới đây đã được đối chiếu lại với mã trước khi sửa.

| # | Mức | Phát hiện | Xử lý |
|---|---|---|---|
| 1 | Cao | Kết quả `ppt_check_deck` (`src/core/deckcheck.ts`) gợi ý sửa bằng `ppt_set_layout`, công cụ thuộc nhóm `sap_xep` có thể đang ẩn khi chỉ bật `kiem_bai` | Đã sửa: "đổi bố cục slide sang bố cục có tiêu đề (chế độ Miễn phí: nhóm sap_xep)"; kiểm thử #1 |
| 2 | Trung bình | Câu lỗi khi ghi vào sheet bị bảo vệ (`src/office/excelTrail.ts`, bọc cả công cụ lõi) nêu `excel_protect` (nhóm `lam_sach_khoa`) | Đã sửa: nêu "công cụ khóa ô công thức, action unprotect; chế độ Miễn phí: nhóm lam_sach_khoa"; kiểm thử #2 |
| 3 | Trung bình | "PivotTable" viết liền không khớp trọn từ `pivot` → không bật nhóm | Đã thêm `pivottable`, `pivot table`; kiểm thử |
| 4 | Trung bình | Chưa có kiểm thử cho chuỗi do công cụ trả về lúc chạy | Đã thêm kiểm thử #1, #2 (quét nguồn `deckcheck.ts`, `excelTrail.ts`). Quét rộng hơn mọi chuỗi tiếng Việt có tên công cụ chuyên: các chỗ còn lại là câu lỗi của chính công cụ cùng nhóm, câu hỏi xác nhận (hiện cho người dùng), Tự kiểm tra, hoặc lời nhắc đã được thay khi chạy |
| 5 | Thấp | Thiếu từ khóa: thả xuống, validation, conditional formatting, lên đầu, hoán đổi, đổi vị trí | Đã thêm. Không thêm "loc" (lẫn "Lộc") và "in" (lẫn "in đậm") |
| 6 | Thấp | Bật nhầm `mo_hinh` với "gia đình", "đặt tên sheet" | Đã thu hẹp thành cụm ("ô giả định", "bảng giả định", "đặt tên vùng", "đặt tên ô"…); kiểm thử |
| 7 | Thấp | Câu ghi chú ảnh đính kèm nêu `ppt_insert_image`; lượt sau không có từ khóa thì nhóm `hinh_anh` không bật | Đã sửa: câu ghi chú không nêu tên công cụ; còn ảnh đính kèm trong phiên thì luôn bật `hinh_anh` |
| 8 | Thấp | Dòng `excel_what_if` mất "hỏi người dùng rồi mới ghi" | Đã thêm lại vào `EXCEL_SPECIAL_LINE`. "Undo removes the whole batch", "and activate it": để lại (hành vi do mã bảo đảm) |
| 9 | Thấp | Lề 40 điểm (lời nhắc cũ) và 36pt (chuẩn làm việc) gộp còn 36 | Để lại: chuẩn làm việc là chỗ duy nhất còn số, không mâu thuẫn |
| 10 | Nit | Chú thích ở `turn.ts` nói chuẩn làm việc nêu tên công cụ chuyên | Đã sửa chú thích (giữ lớp phòng thủ) |

Sau sửa: kiểm thử máy 701/701, tsc sạch. Các sửa chỉ đổi chuỗi và từ khóa nên chưa Chạy thật lại trong Office.
