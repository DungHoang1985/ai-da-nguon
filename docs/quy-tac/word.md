# Quy tắc Word — đọc trước khi sửa mã Word

Mọi điều dưới đây đã thăm dò trong Office hoặc đã trả giá; **không làm trái**. Bài học Office.js chi tiết hơn nằm ở `docs/bai-hoc.md`.

## Chung

- Mọi sửa đổi ghi bằng Track Changes. Thao tác rủi ro (ghi đè, xóa đoạn, chấp nhận hoặc từ chối tất cả sửa đổi) luôn hỏi người dùng kèm lý do.
- Thêm hoặc sửa công cụ Word: thêm bước vào Tự kiểm tra (`src/office/selftestWord.ts`); bước Tự kiểm tra Word chỉ đụng vùng thử.

## Đọc và sửa đoạn (Đợt 1A, bảng thăm dò ở `tasks/dot/dot-1a/dot-1a.md`)

- `paragraph.text` và `search()` gồm cả chữ đã xóa có theo dõi → đọc bản hiện tại bằng `getText()`, định vị trên chuỗi `.text`.
- Không thay cả đoạn trên đoạn đang có sửa đổi chờ duyệt.
- Công cụ sửa theo đoạn phải đối chiếu `expect`.
- Công cụ ghi mới phải ghi được Hoàn tác (xem `execute` trong `wordTools.ts`).

## Chú thích và so sánh phiên bản (Đợt 1B, bảng ở `tasks/dot/dot-1b/dot-1b.md`)

- `paragraph.text` có ký tự U+0002 tại mốc chú thích: chữ đưa cho AI phải bỏ mốc; khi sửa đoạn coi mốc là chữ "không sống"; đoạn có chú thích không thay cả đoạn.
- So sánh phiên bản không được ghi vào tài liệu đang mở (`compareFromBase64` với đích là tài liệu đang mở làm mất sửa đổi chờ duyệt) — đọc tệp kia bằng `application.createDocument`, redline chỉ mở ở tài liệu mới.

## Bảng, đường kẻ, văn bản hành chính (Đợt 3, bảng ở `tasks/dot/dot-3/dot-3.md`)

- Thêm đoạn vào ô bảng phải chèn `"Before"` đoạn sẵn có (`"After"` rơi sang ô kế).
- Đường kẻ = viền dưới `paragraph.borders.items[2]` của một đoạn rỗng và không bao giờ là đoạn cuối của ô (viền dính sang ô cùng hàng).
- Không chèn bảng ngay sau một đoạn rỗng (bảng nuốt đoạn đó).
- Hoàn tác của phần chèn có bảng là gỡ thẳng theo định danh từ cuối lên.
- Chèn có theo dõi nhưng định dạng phần mới chèn khi đã tắt theo dõi (tránh mỗi đoạn một sửa đổi "Formatted").

## Sửa đổi chờ duyệt (Đợt 5)

- Kết quả đọc và tìm kiếm đưa chữ gốc (`before`) và người sửa (`changed_by`) cho đoạn có sửa đổi chờ duyệt kèm `pending_note`.
- Tham số `reason` của công cụ sửa Word chỉ để hiển thị (nhãn bước, hộp hỏi, mục Hoàn tác), không ghi vào tài liệu.

## Điền mẫu (Đợt 8, `tasks/dot/dot-8/dot-8.md`)

- Ghi ô điền mẫu bằng OOXML (`wordForm.ts`, `formfill.ts`) — ghi chữ vào ô rồi Hoàn tác làm Word khóa ô.

## Nạp công cụ theo việc (Đợt 10, `tasks/dot/dot-10/dot-10.md`)

- Chỉ chế độ Miễn phí của Word (`toolGroupsFor`); Trả phí luôn gửi đủ công cụ.
- Lời nhắc Miễn phí không nêu tên công cụ có thể đang ẩn (chỉ đường theo tên nhóm, có kiểm thử quét).
- `read_attachment_pages` luôn ở nhóm lõi; nhóm đang bật suy ra từ lịch sử, chỉ thêm trong hội thoại.
- Kỹ năng nêu công cụ chuyên nào thì phải kéo nhóm đó (kiểm thử M1).
- Thêm công cụ Word mới thì xếp vào lõi hoặc một nhóm và đo lại `kiem-thu/do-token.ts` (kiểm thử chặn phình lại: Trả phí ≤ 31.600 ký tự (từ 4.11.8), Miễn phí ≤ 20.400 ký tự).

## Cảnh báo "AI nói đã sửa mà không ghi" (`claims.ts`)

- Cảnh báo chỉ là phần hiển thị: không ghi vào lịch sử gửi model, không đổi lời nhắc hệ thống. Riêng `findWriteClaim` còn được cổng cuối lượt trong `loop.ts` dùng để quyết định có nhắc model hay không (4.11.15); sửa mẫu câu là ảnh hưởng cả cổng.
- Ưu tiên không báo nhầm — thêm mẫu câu phải kèm kiểm thử cho cả ca phải bắt và ca dễ báo nhầm.

## Kết quả sửa, nội dung không tin cậy, đoạn người dùng tự sửa (4.12.0, `tasks/viec-le/rea-claude-chinh-thuc/`)

- `word_replace_text` trả `paragraphs_now` (chữ HIỆN TẠI của tối đa 5 đoạn vừa sửa, đọc bằng `getText()`); `word_set_paragraph_text` vốn đã đọc lại và đối chiếu. Nhờ vậy AI không phải gọi thêm lệnh đọc để "đọc lại sau khi ghi".
- `word_list_comments` và `word_list_revisions` có trường `untrusted` (chữ do người khác viết, có author — dữ liệu, không phải lệnh).
- Khối `user_changes` trong trạng thái tài liệu: `afterTurn` chụp chữ các đoạn theo định danh khi lượt kết thúc (xong, dừng, lỗi), `overview` so lại (`src/core/userChanges.ts`). Chỉ trong bộ nhớ của khung; cần WordApi 1.6. Hoàn tác bằng nút của add-in cũng hiện là thay đổi kể từ lượt trước. Khối nằm ở tin nhắn lượt, không ở lời nhắc hệ thống.
