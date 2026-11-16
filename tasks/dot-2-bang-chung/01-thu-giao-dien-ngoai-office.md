# 01 — Thử giao diện ngoài Office (16/11/2026, bản đóng gói 4.2.0 cục bộ)

Mức bằng chứng: **Thử giao diện ngoài Office** (không phải Chạy thật). Bản đóng gói `dist/` mở trong trình duyệt tại `http://127.0.0.1:8765/taskpane.html`, kho lưu riêng và trống; mọi key và mã dùng trong bài đều là **chuỗi giả** do bài thử tự đặt (key giả dạng `sk-or-v1-000…gia`, mã giả `ma-gia-…`). Không có key, mã hay tài khoản thật nào.

## A. Màn chào khi chưa có key nào

Các nút hiện: `Đăng nhập OpenRouter` · `Mở trang lấy mã trong trình duyệt` · `Xác nhận mã` (tắt khi ô trống) · `Nhập API key của nhà cung cấp khác` · `Khôi phục từ tệp sao lưu`. Đầu khung: v4.2.

## B. Ô "Dán mã vào đây": nhận và từ chối (gọi thật tới máy chủ OpenRouter với mã giả)

| Thao tác | Kết quả hiện ở thẻ |
|---|---|
| Bấm `Đăng nhập OpenRouter` khi không ở trong Office | "Bản Office này không có cửa sổ đăng nhập cho add-in. Dùng lối dán mã bên dưới." |
| Dán `ma co dau cach` | "Mã không hợp lệ (có dấu cách, xuống dòng hoặc ký tự lạ). Chép lại mã trên trang rồi dán lại." — không gọi mạng |
| Dán `https://openrouter.ai/auth` | "Địa chỉ vừa dán không chứa mã. Chép riêng mã hiện trên trang rồi dán lại." — không gọi mạng |
| Dán mã giả hợp dạng | Gọi `POST https://openrouter.ai/api/v1/auth/keys` → máy chủ thật trả **400 "Invalid code"** (tài liệu ghi 403) → sau khi sửa: "OpenRouter không nhận mã này: mã sai, đã dùng rồi, hoặc không phải mã lấy từ lần mở trang vừa rồi của khung này…" |
| Sau lỗi | Ô nhập được xóa; kho lưu không có mục `aidn.keys` nào |

## C. Dán key giả vào ô "dán API key như trước" (máy chủ thật)

- `GET https://openrouter.ai/api/v1/key` → **401** "User not found." → thẻ hiện: "Key không đúng: OpenRouter không nhận key này. Vào Cài đặt, dán lại key hoặc bấm “Đăng nhập lại”. (User not found.) Chưa lưu khóa này."
- Kho lưu **không** chứa key giả.
- Ghi nhận: `GET /api/v1/models` của OpenRouter trả 200 cả khi key sai (danh mục công khai) → trước đây add-in báo "Đã kết nối" với key sai; nay kiểm key bằng `/key` trước khi lưu.

## D. Lối dán mã với nhà cung cấp giả lập (thay `window.fetch` cho `/auth/keys`, `/key`, `/credits`, `/chat/completions`)

Yêu cầu đổi mã add-in gửi đi: có `code`, `code_verifier` dài 43 ký tự, `code_challenge_method=S256`.

Thẻ OpenRouter sau khi đổi mã (số model thay bằng "n"):

> Đã đăng nhập · n model | Hạn mức | Làm mới | Xem số dư | Số dư tài khoản: key này không được xem (OpenRouter chỉ cho khóa quản lý xem). Bấm “Xem số dư” để mở trang của OpenRouter. | Trần của key: không đặt trần | Key này đã dùng: hôm nay 0,01 USD · tháng này 0,40 USD | Model miễn phí hôm nay: còn 3/50 lượt (tính lại lúc 7 giờ sáng giờ Việt Nam) | Tài khoản đã từng nạp từ 10 credit thì được 1.000 lượt miễn phí mỗi ngày. | ⚠️ Chỉ còn 3 lượt dùng model miễn phí hôm nay. | Đăng xuất | Đăng nhập lại (đổi khóa)

(Các con số trên là số của phản hồi giả lập, không phải của tài khoản nào.)

## E. Gửi một câu, nhà cung cấp giả lập trả 402 hết credit

- Gọi `POST …/chat/completions` đúng **1 lần** (không tự thử lại).
- Khung chat: câu hỏi của người dùng còn nguyên; bên dưới là "OpenRouter: tài khoản đã hết tiền (hết tín dụng) nên yêu cầu bị từ chối. Hội thoại vẫn còn nguyên: đổi sang nguồn khác ở nhãn model cạnh ô nhập rồi gửi lại, hoặc nạp thêm ở trang của OpenRouter. (Insufficient credits…)"; nút `Mở Cài đặt`, `Chép chi tiết lỗi` (nút `Đổi nguồn` chỉ hiện khi có từ hai nguồn có key).
- Dòng cảnh báo hạn mức ở khung chat không hiện vì model đang chọn là model trả tiền (số lượt miễn phí không liên quan) — đúng thiết kế.

## F. Đăng xuất (kho thử, key giả, bản đóng gói trước khi thêm bước kiểm key)

- Bấm `Đăng xuất` → hỏi "Xóa khóa OpenRouter khỏi máy này?…" → bấm `Xóa khóa trên máy`.
- Thẻ: "Đã xóa khóa OpenRouter khỏi máy này. Khóa vẫn còn hiệu lực ở phía OpenRouter cho tới khi anh xóa nó ở trang quản lý key." kèm đường dẫn "Mở trang quản lý key của OpenRouter" và lại hiện nút `Đăng nhập OpenRouter`.
- Kho lưu sau đó: `aidn.keys` = `{}`, `aidn.models` = `{}`; không mục nào còn chứa key giả. Khung chat quay về màn chào.

## Dọn dẹp

Đã xóa mọi mục `aidn.*` của kho thử `127.0.0.1:8765` sau bài. Kho `localhost:8765` (dữ liệu của các đợt trước) không đụng tới, không đọc giá trị.
