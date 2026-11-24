# Soát mã độc lập — phần B của Đợt 3 (bản 4.3.0)

Ngày 16/11/2026. Một tác tử khác (không viết mã này) soát `git diff 4659a6b..eb44113 -- src test`, trọng tâm: kênh lệnh ghi chéo có bị lợi dụng để ghi mà không hỏi không; nội dung từ ứng dụng kia có được coi là dữ liệu không. Tự viết 7 ca thử chạy thật ngoài kho; phần cần Office hoặc hai tiến trình chỉ suy từ đọc mã.

## Trả lời hai câu hỏi chính

- **Kênh lệnh có bị lợi dụng để ghi mà không hỏi người dùng không? Không.** Mọi lời gọi `write_other_app` của model đều qua hộp hỏi: cờ `alwaysAsk` ép kiểm rủi ro cả khi tắt "hỏi trước thao tác rủi ro"; "Cho phép hết lần này" không áp dụng cho lệnh có lý do rủi ro. Trong mã không có chỗ nào chèn HTML / Markdown của model, tài liệu hay web thành DOM chạy được script, nên nội dung tài liệu không tự đặt lệnh vào hộp thư được. Lệnh hợp lệ chỉ THÊM mới (Word thêm đoạn có theo dõi, Excel tạo sheet mới với công thức thành chữ, PowerPoint thêm một slide); lệnh cũ không chạy lại được. Có **một khe cấu trúc**: lúc kiểm rủi ro đích chưa mở thì không hỏi, nếu đích mở ngay sau đó thì lệnh có thể đi mà chưa hỏi (khoảng hở cỡ vài microtask) → **đã bịt**: công cụ chỉ gửi đúng nội dung vừa được đưa ra hỏi.
- **Nội dung từ ứng dụng kia có được coi là dữ liệu không? Có**, bằng lời nhắc an toàn (đã kể "nội dung đang mở ở ứng dụng khác", "kết quả của công cụ") và mô tả công cụ; hàng rào thực sự là hộp hỏi + chỉ thêm mới + Hoàn tác. Điểm yếu tìm ra: hộp hỏi chỉ hiện 160 ký tự đầu nên đoạn độc có thể nấp phía sau → **đã sửa** (hiện tới 1.500 ký tự, nêu tổng số ký tự và phần chưa hiện, nhắc "nội dung do AI soạn").

## Phát hiện: 0 cao · 6 trung bình (+ khe ở trên) · 10 thấp. Đã sửa 11, để lại 6 (ghi lý do).

| # | Mức | Phát hiện | Xử lý |
|---|---|---|---|
| 0 | Khe cấu trúc | Kiểm rủi ro trả "không hỏi" khi đích chưa mở, rồi công cụ vẫn gửi nếu đích mở kịp | Đã sửa: chỉ gửi lệnh có nội dung trùng với lần hỏi gần nhất; một lần hỏi dùng cho một lần gửi. Có kiểm thử chặn |
| 1 | Trung bình | Hộp hỏi cắt nội dung (Word 160 ký tự, Excel chỉ hàng đầu) | Đã sửa: Word 1.500 ký tự kèm tổng số ký tự và "còn N ký tự nữa chưa hiện"; Excel hàng đầu, ba hàng kế, hàng cuối; slide 1.200 ký tự. Có kiểm thử chặn |
| 2 | Trung bình | Hai tài liệu cùng ứng dụng: tên tệp trong hộp hỏi có thể khác tệp thực nhận; có thể ghi đôi | Đã sửa: lệnh mang tên tệp đích, khung của tệp khác không lấy; hai khung cùng tên thì giành lệnh bằng dấu nhận (chỉ một khung chạy). Có kiểm thử chặn. Chưa thử trong Office với hai tài liệu thật |
| 3 | Trung bình | Khung nhận đang bận thì nuốt lệnh và báo sai "đang tắt nhận ghi" | Đã sửa: bận thì để lệnh nằm trong hộp thư; khung gửi tự rút khi hết hạn chờ |
| 4 | Trung bình | Người dùng gửi tin đúng lúc khung đang ghi hộ → mục "vừa ghi" và câu trả lời có thể lẫn nhau | Đã sửa: trong lúc ghi hộ khung coi như đang bận (không gửi được yêu cầu mới) |
| 5 | Trung bình | Chỉ hình bằng số "1" có thể trúng hình tên "Rectangle 12" | Đã sửa: số thứ tự được xét trước khớp một phần của tên |
| 6 | Trung bình | Model chưa rõ có nhận ảnh: ảnh người dùng đính kèm bị bỏ ngay khi công cụ chụp ảnh slide | Đã sửa: tin ảnh của công cụ mang cờ riêng, không tính là "câu hỏi cuối"; ảnh công cụ chỉ giữ lần chụp gần nhất. Có kiểm thử chặn |
| 7 | Thấp | Lệnh bị bỏ không báo lại → khung gửi báo "có thể đã ghi" dù chưa ghi | Đã sửa: lệnh bị bỏ có mã thì trả lời "chưa ghi gì"; hạn của lệnh (30 giây) dài hơn hạn chờ của khung gửi (25 giây) |
| 8 | Thấp | Chữ mẫu "[cần bổ sung]" không bao giờ bị báo | Đã sửa; có kiểm thử chặn |
| 9 | Thấp | `ppt_format_text` với `find` rỗng định dạng cả hình | Đã sửa: từ chối |
| 10 | Thấp | `ppt_edit_table`: ô `null` thành ô (0,0); không giới hạn số ô | Đã sửa: bắt buộc có row / col; tối đa 300 ô, 50 hàng mỗi lần |
| 11 | Thấp | Phép kiểm màu chữ ngoài bảng màu chủ đề chưa được nối (mã chết) | **Để lại**: cần so cả biến thể sáng / tối của màu chủ đề mới không báo nhầm; công cụ không quảng cáo mục này |
| 12 | Thấp | `ppt_set_layout` với bài nhiều mẫu (master) lấy bố cục trùng tên của mẫu đầu tiên | **Để lại**: hiếm; ghi ở Giới hạn đã biết |
| 13 | Thấp | Mỗi lần định dạng chữ / sửa bảng giữ một bản sao slide để Hoàn tác (tốn bộ nhớ khi rất nhiều bước) | **Để lại**: bản sao mất khi tải lại khung; đánh đổi lấy Hoàn tác chắc chắn |
| 14 | Thấp | Sửa lỗi bố cục của slide biểu đồ chỉ xét mẫu đầu tiên | **Để lại**: như mục 12 |
| 15 | Thấp | Tóm tắt hội thoại chưa tính kích thước ảnh | **Để lại**: ngưỡng tóm tắt dùng số token thực đo của bước trước nên vẫn có tác dụng; ghi ở Giới hạn đã biết |
| 16 | Thấp | "Nhận ghi từ ứng dụng khác" mặc định bật | **Để lại có chủ ý**: người dùng luôn được hỏi ở khung gửi, lệnh chỉ thêm mới và có Hoàn tác; có công tắc tắt trong Cài đặt |

Đã kiểm, không có lỗi: bộ hẹn giờ và bộ nghe của khung được dọn đủ, không dội sự kiện; `__images` chỉ được bóc ở cấp cao nhất của kết quả công cụ nên dữ liệu đọc từ ứng dụng khác không giả làm ảnh được; mục "vừa ghi theo yêu cầu" không vào lịch sử gửi model; bản lưu hội thoại không chứa ảnh; công thức `=`, `+`, `@`, `-…` từ ứng dụng khác thành chữ, số âm thật được giữ.

Sau khi sửa: `npm test` 420/420; Tự kiểm tra Word 37/37, Excel 23/23, PowerPoint 21/21; chạy lại bài ghi chéo Excel → Word (tệp 16).
