# Soát mã độc lập — mã từ 4.11.0 tới 4.11.4 (22/11/2026)

Phạm vi: `git diff 8546999..d1b4eb9 -- src` (bảng PowerPoint tự co, cảnh báo qua Làm tiếp, đo token đầu vào, ngày trong Excel, `more`, kiểm toán mô hình, thứ tự xoay vòng Miễn phí). Một tác tử soát riêng, không thấy kết luận của phiên làm; nó chạy `npm test` (731 đạt) và `tsc` (sạch), không sửa tệp.

## Đã sửa

| Mức | Mục | Lỗi | Sửa | Kiểm |
|---|---|---|---|---|
| Cao | C1 | `tableColumnWidths`: từ 13 cột trở lên, sàn 8% × số cột vượt 100% → cột cuối ra bề ngang âm, `ppt_add_table` hỏng (có từ 4.11.1) | Sàn không quá phần chia đều; sàn chiếm hết thì chia đều | Kiểm thử máy 12, 13, 14, 20, 40 cột: mọi cột > 0, tổng đúng |
| Cao | C2 | `excel_format_range` `more` lọt qua lớp chặn sheet "Nhật ký AI" (có từ 4.11.4) | `targetsLogSheet` xét cả từng mục của `more` | Kiểm thử máy: mục nhắm nhật ký bị chặn, mục thường không |
| TB | M1 | Lời nhắc "còn dữ liệu cũ bên dưới" báo ở gần như mọi lần sửa ô giữa bảng và thúc AI xóa (`excel_clear_range`) | Chỉ nhắc khi ghi đè cả bảng từ đầu bảng (≥ 2 hàng, vùng cũ kín, hàng trên trống); câu nhắc trung tính, không ra lệnh xóa | Tự kiểm tra Excel: ghi đè bảng ngắn hơn → nhắc; sửa một ô giữa bảng → không nhắc |
| TB | M2 | Hệ số Claude 1,5 ký tự / token đo trên tiếng Việt → tệp tiếng Anh dài bị chặn nhầm ở Trả phí | Hệ số chữ của Claude theo tỷ lệ ký tự ngoài ASCII: 3,8 (tiếng Anh) → 1,5 (tiếng Việt) | Kiểm thử máy hai ca |
| TB | M3 | Với Claude, ngưỡng chặn "không chứa nổi" đến trước ngưỡng tự tóm tắt → hội thoại dài bị chặn thay vì được tóm tắt | Tự tóm tắt còn kích hoạt khi số ước token của lịch sử ≥ 60% sức chứa (cùng bộ ước với lượt gửi) | Kiểm thử máy cũ (tóm tắt) vẫn đạt |
| Thấp | L2 | Định dạng ngày ghi lại cả ma trận công thức (có thể đụng mảng động, mảng CSE) | Chỉ ghi từng ô đã đổi | Tự kiểm tra Excel 42/42 |
| Thấp | L3 | Hoàn tác trả "@" trước khi ghi → ô công thức có định dạng "@" thành chữ | Ô công thức "@": trả General, ghi, rồi đặt lại "@" | Tự kiểm tra Excel 42/42 |
| Thấp | L6 | Luật "ô cuối dãy" bỏ hết chữ số nên báo nhầm ô cuối khác hằng số (tăng giá năm cuối `*1.05`) | Chỉ bỏ số trong địa chỉ R1C1 | Kiểm thử máy: ca dễ báo nhầm không báo, ca SUM hụt vẫn báo |

## Để lại (có lý do)

| Mục | Lý do |
|---|---|
| L1 `ppt_edit_table` làm phẳng cỡ chữ cả bảng khi bảng tràn | Chỉ chạy khi bảng tràn đáy slide (không tràn thì không đụng), có Hoàn tác bằng bản sao slide; co theo tỷ lệ từng ô để đợt sau |
| L4 Chữ như "1.2.2024", "12-10-2023" ghi vào ô thường thành ngày | Đánh đổi có chủ ý: Excel cũng tự đổi các chữ này thành ngày (nhưng theo kiểu Mỹ); add-in đổi theo ngày/tháng. Ô định dạng chữ "@" thì giữ nguyên |
| L5 Gộp lệnh ghi qua Làm tiếp có thể làm mất cảnh báo đúng | Đúng nguyên tắc "ưu tiên không báo nhầm" của `claims.ts` |

Kiểm sau khi sửa: kiểm thử máy 734/734, kiểm kiểu sạch; Tự kiểm tra trên bản cục bộ (desktop ẩn): Excel 42/42 (`tu-kiem-tra-excel-soat-ma-cuc-bo.json`), PowerPoint 26/26 (`tu-kiem-tra-ppt-soat-ma-cuc-bo.json`). Chưa thử trong PowerPoint một bảng 14 cột thật.
