# Quy tắc vòng tác tử, tệp đính kèm, bộ nhớ, ghi chéo — đọc trước khi sửa `src/core/agent/`, `pdfPages.ts`, `memory.ts`, `crossapp.ts`, `crossExec.ts`, `skillpack.ts`

**Không làm trái.**

## Ảnh do công cụ chụp

- `__images` trong kết quả công cụ được vòng lặp tách ra, gửi ở một tin nhắn user riêng có cờ `fromTool` và câu "không phải yêu cầu mới của người dùng".
- Không nhét ảnh vào kết quả dạng chữ, không lưu ảnh xuống kho.

## Tệp đính kèm theo trang (Đợt 5, `tasks/dot/dot-5/dot-5.md`)

- Ảnh trang quét chỉ gửi qua kênh ảnh công cụ (`__images`), tối đa 4 trang một lần gọi.
- Lời dặn chép nguyên văn nằm trong kết quả công cụ, không vào lời nhắc hệ thống.
- Nguồn không nhận PDF mà danh mục không ghi rõ model nhận ảnh thì KHÔNG gửi ảnh trang (`canSeePages` — gửi là bị từ chối cả lượt).
- PDF quét từ 5 trang trở xuống với nguồn đọc được nguyên tệp vẫn gửi nguyên tệp.
- `read_attachment_pages` luôn có trong bộ công cụ chung (không thêm bớt theo lượt).

## Bộ nhớ dài hạn (Đợt 12 E, `src/core/memory.ts`)

- Khối bộ nhớ chụp MỘT lần khi hội thoại bắt đầu và lưu kèm hội thoại — không dựng lại theo lượt (giữ bộ nhớ đệm).
- AI chỉ đề xuất (`propose_memory`, chỉ nhận khi câu người dùng gõ ở lượt đó nói tới ghi nhớ — chặn chỉ dẫn ẩn trong tài liệu), chỉ lưu khi người dùng bấm Nhớ.
- Mỗi lượt dựng khối từ bản chụp theo trần của chế độ đang dùng; mục đã xóa / tắt thôi gửi ngay.
- Chữ mục nhớ bỏ ký tự vô hình.
- Bộ lọc chặn key, mật khẩu, số điện thoại, số định danh, số tài khoản (thêm mẫu phải kèm ca dễ chặn nhầm).
- Chế độ Miễn phí trần 800 ký tự và `propose_memory` nằm trong nhóm `bo_nho` nạp theo từ khóa (sức chứa model miễn phí nhỏ — anh dặn 23/11/2026).

## Ghi chéo ứng dụng (Đợt 3, `tasks/dot/dot-3/dot-3.md`)

- `write_other_app` luôn hỏi người dùng ở khung gửi (cờ `alwaysAsk`; "Cho phép hết lần này" và việc tắt "hỏi trước thao tác rủi ro" đều không bỏ qua được) và chỉ gửi đúng nội dung vừa hỏi.
- Hộp hỏi phải hiện gần đủ nội dung và tên tệp đích.
- Khung nhận chỉ chạy các thao tác THÊM mới trong danh sách đóng của `crossapp.ts`, kiểm lại lệnh, ghi bằng công cụ của chính nó (Word: Track Changes) và để lại mục có nút Hoàn tác.
- Nội dung lệnh không bao giờ đưa cho model ở khung nhận; công thức từ ứng dụng khác ghi thành chữ.
- Thêm thao tác ghi chéo mới phải qua soát mã độc lập.

## Kỹ năng dựng sẵn

- Không ghi số hiệu văn bản pháp luật (có kiểm thử chặn).
