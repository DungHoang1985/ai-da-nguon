# Quy tắc PowerPoint — đọc trước khi sửa mã PowerPoint

Đã thăm dò trong Office; **không làm trái**.

## Bảng (thăm dò 22/11/2026, `tasks/viec-le/sua-22-11/sua-22-11.md`)

- Cỡ chữ bảng đặt bằng `cell.font.size` từng ô SAU khi thêm (PowerPointApi 1.9) — `uniformCellProperties` làm mất kiểu bảng.
- Không gán tên phông cho ô bảng (để theo chủ đề của mẫu).
- Chiều cao thật của bảng là `shape.height` (đọc ngay sau `sync`); `rowHeight` chỉ là mức tối thiểu.
- `ppt_add_table` tự co chữ tới 10 pt cho vừa slide và báo `box`, `fits`.

## Xóa và ghi đè (4.11.13)

- Xóa slide / hình chỉ khi đã xuất được slide để Hoàn tác (`exportSlide`); không xuất được thì dừng trước lệnh `delete()` (có kiểm thử chặn).
- Ghi đè chữ (anh chọn 25/11/2026): chỉ hỏi khi THAY hoặc XÓA chữ dài đang có — hình trên 200 ký tự, ô bảng trên 100 ký tự (`longOverwrite` trong `src/office/util.ts`); chữ ngắn, hoặc chữ mới vẫn chứa nguyên chữ cũ, thì không hỏi.

## Giới hạn API (Đợt 3, bảng ở `tasks/dot/dot-3/dot-3.md`)

- Không có API cho ghi chú diễn giả, số liệu của biểu đồ đã có, phông chủ đề (đọc từ gói của slide xuất ra), xóa bố cục.
- Hình trùng tên phải chỉ bằng số thứ tự.
- Số đo đưa cho AI phải sát với cái mắt thấy (so vùng chữ, không so khung rỗng).

## Xem lại sau khi sửa bố cục (4.12.0, `tasks/viec-le/rea-claude-chinh-thuc/`)

- Lệnh làm đổi bố cục mang cờ `needsLook` (danh sách `PPT_NEEDS_LOOK` trong `pptTools.ts`); `ppt_get_slide_image` và `ppt_check_deck` mang cờ `look`. Thêm công cụ ghi đổi bố cục mới thì thêm vào danh sách (kiểm thử chặn tên lạ).
- Danh sách lỗi khách quan khi nhìn ảnh (`SLIDE_CHECKLIST`, `pptDesign.ts`) đi trong KẾT QUẢ công cụ, không trong lời nhắc hệ thống (giữ trần ký tự). Chỉ lỗi khách quan, không chấm thẩm mỹ; tối đa 3 vòng sửa cho cùng một lỗi rồi hỏi.
