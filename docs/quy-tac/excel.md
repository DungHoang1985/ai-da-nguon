# Quy tắc Excel — đọc trước khi sửa mã Excel

Đã thăm dò trong Office; **không làm trái**. Bảng thăm dò ở `tasks/dot/dot-4/dot-4.md` và `tasks/dot/dot-12/dot-12.md`.

## Ngày (Đợt 12, thăm dò 22/11/2026)

- Office.js đọc chữ "05/03/2026" ghi vào ô thường theo kiểu Mỹ (tháng/ngày) — công cụ ghi phải tự đổi chữ dạng ngày thành số ngày theo ngày/tháng/năm (`src/core/datetext.ts`).
- "dd/mm/yyyy" trùng kiểu ngày ngắn của Windows thì Excel lưu thành định dạng dựng sẵn (đọc ra "m/d/yyyy") → đặt "dd/mm/yyyy;@".
- Hoàn tác trả định dạng số TRƯỚC rồi mới ghi nội dung.

## Nhật ký và bộ công cụ (Đợt 4)

- Nhật ký và tô sáng **mặc định tắt**; lúc bật phải có câu cảnh báo sheet đi theo tệp.
- Mọi chữ ghi vào sheet "Nhật ký AI" phải qua `logText` (không bao giờ thành công thức).
- AI không ghi, sửa, đổi tên sheet nhật ký (lớp chặn so tên SAU khi làm sạch và tính cả lệnh không nêu sheet); chỉ được xóa cả sheet khi người dùng nhờ.
- Nhật ký không vào lịch sử gửi model.
- Bộ công cụ Excel chỉ dùng qua `getHostTools` (đã bọc `withAuditTrail`, gỡ `__trail` khỏi kết quả).
- `excel_what_if` luôn trả ô đầu vào về nội dung gốc và không tự ghi kết quả.
- `excel_protect` không bao giờ đặt mật khẩu.
- Phép kiểm rủi ro của công cụ ghi phải chạy trên dữ liệu đã chuẩn hóa bằng đúng hàm công cụ dùng (`normalize2D`).
- Tự kiểm tra Excel chạy với nhật ký tắt và sheet nhật ký thử riêng.

## Câu hỏi trước thao tác ghi (4.11.13)

- Câu hỏi chỉ nói "(có thể Hoàn tác)" khi vùng không quá 10.000 ô (`MAX_UNDO_CELLS`, `undoNote` trong `src/office/util.ts` — cùng ngưỡng với `snapshot`); vùng lớn hơn phải nói KHÔNG hoàn tác được, và kết quả công cụ thêm `noUndoSuffix`.
- Sắp xếp vùng trên 10.000 ô phải hỏi (vùng nhỏ có Hoàn tác thì không hỏi).
- Không đọc được vùng đích để kiểm ghi đè / đếm ô / đọc bố trí khóa thì **hỏi** (`OVERWRITE_UNCHECKED`), không `.catch(() => null)` cho chạy im lặng.

## Kiểm toán mô hình và công thức cần hỏi

- `modelaudit.ts` và `formularisk.ts` ưu tiên không báo nhầm: thêm mẫu phải kèm kiểm thử cho cả ca phải bắt và ca dễ báo nhầm (tham chiếu bảng `[2024]`, hàng số nhập tay xen công thức, gạch đầu dòng "- …").

## Câu hỏi ghi đè (4.12.0)

- Câu hỏi trước khi ghi đè nêu tối đa 5 ô sẽ mất kèm giá trị / công thức đang có (`excelRisk`, `excelAdvanced.ts`); Tự kiểm tra có bước kiểm dạng câu này.
