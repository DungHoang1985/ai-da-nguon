# Quy tắc nguồn AI, chế độ, đăng nhập — đọc trước khi sửa `providers.ts`, `llm.ts`, `modes.ts`, `rotation.ts`, `oauth.ts`, `inputsize.ts`, `catalog.ts`

**Không làm trái.** Tài liệu của nhà cung cấp đổi nhanh: đối chiếu tài liệu chính thức hiện hành trước khi viết mã gọi API.

## Nguồn nói kiểu Anthropic (Đợt 6, `tasks/dot/dot-6/dot-6.md`)

- Yêu cầu gửi `api.anthropic.com` không được đổi (có kiểm thử chặn).
- Khác biệt của từng nguồn chỉ khai trong `dialect`, không rẽ nhánh theo mã nguồn trong `llm.ts`.
- Nguồn `bare` chỉ gửi tiêu đề key và `content-type` (thêm `anthropic-version` là bị trình duyệt chặn ở MiniMax, Kimi, OpenRouter).
- Nguồn `keyFrom` không lưu bản sao key.
- Nguồn tự khai chỉ nhận `https`, mã phải có tiền tố `tuy-chinh-`, không đè được nguồn có sẵn, và thẻ phải hiện tên miền nhận key.
- Thêm sẵn một nguồn mới chỉ sau khi đọc tài liệu chính thức và đo bằng trình duyệt thật (Edge `--headless=new`, cả bước dò trước với đúng bộ tiêu đề).

## Chế độ Miễn phí / Trả phí (Đợt 9, `tasks/dot/dot-9/dot-9.md`)

- Chế độ Miễn phí không bao giờ gọi nguồn hoặc model ngoài danh sách miễn phí (`modelInMode`), kể cả khi hết mọi nguồn.
- Mọi lần gọi model của lượt (tóm tắt, soát lập luận, tác tử, trò chuyện) đi qua `FreeRotation`.
- Không gửi plugin tìm web của OpenRouter (tính tiền) ở chế độ Miễn phí.
- Nghỉ có hạn (lỗi tạm thời, hết lượt ngày / tháng) khác bỏ cả lượt (key, quá lớn, không gọi được công cụ).
- Gemini nhận lệnh gọi công cụ của nguồn khác dưới dạng chữ (`foreignToolsAsText`) — chỉ đổi yêu cầu gửi đi, không đổi lịch sử.
- Chế độ Trả phí giữ nguyên cách thử lại cũ, không xoay vòng.

## Đo token đầu vào (Đợt 12, `src/core/inputsize.ts`)

- Đo đủ lời nhắc hệ thống, bộ công cụ, lịch sử và tệp.
- Hệ số ký tự / token theo họ model chỉ đổi khi có số đo Chạy thật (`usage` của lần gửi đầu, thuộc tính `data-input-est` / `data-first-input` trên dòng token) — bảng ở `tasks/dot/dot-12/dot-12.md`.
- Chế độ Trả phí chỉ báo và gợi ý, không tự đổi model.
- Bài thử "quá lớn" phải chọn model có sức chứa đã biết nhỏ hơn số đo (lần đầu chọn nhầm Haiku 1.000.000 token và gửi thật).

## Model mặc định, khối suy nghĩ

- Model mặc định và dấu ⚠ "đã thấy không gọi công cụ" (`catalog.ts`) chỉ đổi khi có bằng chứng Chạy thật và anh đã chọn.
- Khối suy nghĩ của model (Anthropic `thinking`, OpenRouter `reasoning_details`) nằm trong `raw` của tin nhắn trợ lý và được gửi lại nguyên văn, đúng thứ tự; chỗ nào buộc phải đổi tin nhắn cũ (tóm tắt, rút gọn khi lưu) thì bỏ khối suy nghĩ đi kèm.
- Tham số suy nghĩ không được nằm trong lời nhắc hệ thống, bộ công cụ hay tin nhắn.

## Đăng nhập và khóa (Đợt 2, `tasks/dot/dot-2/dot-2.md`)

- Mã kiểm của luồng đăng nhập (`code_verifier`) chỉ nằm trong bộ nhớ của khung, không ghi kho lưu, không vào địa chỉ trang.
- Trang hộp thoại `dang-nhap.html` phải cùng tên miền với khung và chỉ chuyển tới tên miền khai trong cấu hình `oauth` (không nhận địa chỉ đích từ tham số).
- Khung chỉ nhận mã khi `state` khớp lần đăng nhập đang mở.
- Mọi câu lỗi của nhà cung cấp đưa ra ngoài (hiện, chép, lưu cùng hội thoại) phải qua `maskSecrets`; `maskSecrets` che cả mã tổ chức / dự án của tài khoản (`org_…`, `proj_…`).
- Key OpenRouter kiểm bằng `GET /key` trước khi lưu (`/models` trả 200 cả khi key sai).
- Lỗi đã rõ lý do về key hoặc tiền (`ApiError.reason`) không tự thử lại.
- Bài thử không đọc kho lưu của khung, không bấm nút mở trình duyệt ngoài khi chưa báo anh, và che số trong thẻ hạn mức trước khi ghi bằng chứng.
