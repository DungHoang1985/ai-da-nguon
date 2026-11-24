# Quy tắc PowerPoint — đọc trước khi sửa mã PowerPoint

Đã thăm dò trong Office; **không làm trái**.

## Bảng (thăm dò 22/11/2026, `tasks/viec-le/sua-22-11/sua-22-11.md`)

- Cỡ chữ bảng đặt bằng `cell.font.size` từng ô SAU khi thêm (PowerPointApi 1.9) — `uniformCellProperties` làm mất kiểu bảng.
- Không gán tên phông cho ô bảng (để theo chủ đề của mẫu).
- Chiều cao thật của bảng là `shape.height` (đọc ngay sau `sync`); `rowHeight` chỉ là mức tối thiểu.
- `ppt_add_table` tự co chữ tới 10 pt cho vừa slide và báo `box`, `fits`.

## Giới hạn API (Đợt 3, bảng ở `tasks/dot/dot-3/dot-3.md`)

- Không có API cho ghi chú diễn giả, số liệu của biểu đồ đã có, phông chủ đề (đọc từ gói của slide xuất ra), xóa bố cục.
- Hình trùng tên phải chỉ bằng số thứ tự.
- Số đo đưa cho AI phải sát với cái mắt thấy (so vùng chữ, không so khung rỗng).
