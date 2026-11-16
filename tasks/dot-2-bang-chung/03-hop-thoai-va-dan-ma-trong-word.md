# 03 — Hộp thoại đăng nhập và lối dán mã trong Word thật (16/11/2026)

Word 16.0 trên máy anh, tệp thử `aidn-thu-dot2.docx`, khung add-in nạp **bản cục bộ 4.2.0** bằng `node kiem-thu/cdp.mjs serve 9229 dist` (đầu khung hiện v4.2). Điều khiển bằng `kiem-thu/cdp.mjs`, không điều khiển màn hình. Thử **tới ngay trước bước anh phải đăng nhập**: không đăng nhập, không tạo khóa, không đổi khóa đang lưu. Mọi mã dùng trong bài là mã giả (`ma-gia-…`). Giá trị nhạy cảm trong địa chỉ chỉ ghi độ dài; số trong thẻ hạn mức thay bằng `#`.

Bài thử: `kiem-thu/cdp-bai/dot-2-trang-thai.js`, `dot-2-mo-hop-thoai.js`, `dot-2-doc-hop-thoai.js`, `dot-2-gia-lap-chuyen-ve.js`, `dot-2-doc-the.js`, `dot-2-dan-ma.js`.

## A. Trạng thái đầu và thẻ hạn mức (key OpenRouter đang lưu của anh, gọi thật)

```
phienBan: v4.2
hopThoaiHoTro (Office.context.ui.displayDialogAsync): true
moTrinhDuyetHoTro (Office.context.ui.openBrowserWindow): true
theOpenRouter: OpenRouter | Lấy key | Nhiều model miễn phí (đuôi :free) | Đã đăng nhập · # model | Hạn mức | Làm mới | Xem số dư |
  Số dư tài khoản: # USD | Trần của key: không đặt trần | Key này đã dùng: hôm nay # USD · tháng này # USD |
  Model miễn phí hôm nay: còn #/# lượt (tính lại lúc # giờ sáng giờ Việt Nam) |
  Tài khoản đã từng nạp từ # credit thì được # lượt miễn phí mỗi ngày. | Đăng xuất | Đăng nhập lại (đổi khóa)
```

Ghi nhận: thẻ có dòng "Số dư tài khoản" → với key đang lưu, `GET /api/v1/credits` **trả được số dư** (tài liệu ghi cần khóa quản lý). Add-in xử lý cả hai trường hợp: xem được thì hiện, không được thì hiện dòng "key này không được xem" kèm đường dẫn (bằng chứng 01, mục D).

## B. Hộp thoại của Office: mở được, tới đúng trang, đúng tham số

Bấm `Đăng nhập lại (đổi khóa)` → `Đăng nhập lại`. Thẻ: "Đang mở cửa sổ đăng nhập OpenRouter… Đăng nhập (nếu được hỏi) rồi bấm Authorize trong cửa sổ đó."

Nhật ký `serve`: hộp thoại là một trang mới, yêu cầu đầu tiên của nó là `dang-nhap.html` (trả bản cục bộ) rồi `dang-nhap.<mã băm>.js`.

Các trang đang mở trên cổng gỡ lỗi sau 5 giây:

```
page  https://openrouter.ai/sign-up?redirect_url=<258 ký tự>
page  https://dunghoang1985.github.io/ai-da-nguon/taskpane.html?_host_Info=<27 ký tự>
```

Đọc trong hộp thoại (OpenRouter đưa người chưa đăng nhập sang trang đăng nhập, giữ địa chỉ gốc trong `redirect_url`):

```
tenMien: openrouter.ai   duongDan: /sign-up   tieuDe: "Sign Up | OpenRouter"
diaChiGoc: openrouter.ai/auth với
  callback_url=https://dunghoang1985.github.io/ai-da-nguon/dang-nhap.html
  code_challenge=<43 ký tự>     (đúng dạng base64url 43 ký tự)
  code_challenge_method=S256
  key_label=AI Da Nguon (Office)
  state=<22 ký tự>
chữ trên trang: Sign Up … First name … Email address … Password … Continue … Already have an account? Sign in
```

Ảnh hộp thoại: `02-hop-thoai-office-trang-openrouter.png` (trang đăng ký trống của OpenRouter, có nút GitHub, Google, MetaMask; không có thông tin cá nhân).

**Kết luận thăm dò (mục 8 của kế hoạch): hộp thoại của Office chạy được luồng này trên máy.** Lưu ý cho anh: trang mở ra là "Sign Up"; đã có tài khoản thì bấm "Sign in" ở cuối trang.

## C. Đường về: nhà cung cấp chuyển về trang của add-in → gửi mã về khung

Giả lập bước chuyển về bằng cách cho hộp thoại đi tới `callback_url?code=<mã giả>&state=…` (state lấy từ chính địa chỉ của hộp thoại, không in ra).

| Ca | Kết quả ở thẻ OpenRouter | Hộp thoại | Khóa đang lưu |
|---|---|---|---|
| state **sai** | "Mã trả về không thuộc lần đăng nhập đang mở (state không khớp). Bấm đăng nhập lại." | tự đóng (danh sách trang chỉ còn khung add-in) | không đổi (thẻ vẫn "Hạn mức", còn nút Đăng xuất) |
| state **đúng**, mã giả | "OpenRouter không nhận mã này: mã sai, đã dùng rồi, hoặc không phải mã lấy từ lần mở trang vừa rồi của khung này. Bấm đăng nhập lại để lấy mã mới (Invalid code)." | tự đóng | không đổi |

Hai ca này chứng minh: trang `dang-nhap.html` vẫn khởi động được office.js sau khi cửa sổ đi qua tên miền khác, `messageParent` tới được khung, khung đối chiếu state, đóng hộp thoại, gọi đổi mã thật tới OpenRouter và dịch đúng lỗi.

## D. Lối dán mã trong Word

| Thao tác | Kết quả |
|---|---|
| Ô trống | nút `Xác nhận mã` tắt |
| Dán `ma co dau cach` | "Mã không hợp lệ (có dấu cách, xuống dòng hoặc ký tự lạ). Chép lại mã trên trang rồi dán lại." |
| Dán `https://openrouter.ai/auth` | "Địa chỉ vừa dán không chứa mã. Chép riêng mã hiện trên trang rồi dán lại." |
| Dán mã giả hợp dạng | "OpenRouter không nhận mã này: … (Invalid code)."; ô nhập được xóa |
| Sau cùng | thẻ vẫn ở trạng thái đã đăng nhập; nhãn model cạnh ô nhập vẫn là model OpenRouter như trước |

Nút `Mở trang lấy mã trong trình duyệt`: lần chạy đầu của bài đã **bấm thật một lần** (hàm thay tạm không có tác dụng vì `Office.context.ui.openBrowserWindow` không ghi đè được: `writable = false`). Kết quả: thẻ báo "Đã mở trang OpenRouter trong trình duyệt…", và trên máy có đúng 1 cửa sổ trình duyệt có chữ "OpenRouter" trong tiêu đề (chỉ đếm, không đọc nội dung). Tức là nút mở được trình duyệt mặc định. Thẻ trình duyệt đó để lại trên máy anh, đóng đi là xong.

Lỗi tìm ra nhờ bài này và đã sửa: sau khi bấm `Mở trang lấy mã…`, nút `Xác nhận mã` bị tắt (dòng hướng dẫn dùng nhầm trạng thái "đang bận"). Các dòng ở bảng trên là của lần chạy lại trên bản đã sửa.

## Chạy lại sau khi sửa theo lần soát mã độc lập (16/11/2026, commit `e352118`)

Cùng các bài, trên bản đóng gói mới: `08a-chay-lai-trang-thai.json` (thẻ hạn mức, số được che), `08b-chay-lai-hop-thoai.json` (hộp thoại tới `openrouter.ai/auth` qua trang đăng nhập, đúng tham số), `08c-chay-lai-duong-ve.json` (state đúng + mã giả → "OpenRouter không nhận mã này…", khóa đang lưu không đổi), `08d-chay-lai-dan-ma.json` (thêm ca dán nhầm chuỗi có dạng API key của nhà cung cấp khác: bị từ chối tại chỗ "Chuỗi vừa dán là API key, không phải mã đăng nhập OpenRouter…", **0 lần gọi mạng**). Kết quả không đổi so với các bảng trên.

## Chưa kiểm được ở bài này (cần anh ngồi trước máy)

Đăng nhập thật trong hộp thoại, bấm Authorize, nhận khóa, danh sách model hiện, sửa một đoạn Word (tiêu chí 1); chép mã thật từ trang rồi dán (tiêu chí 2); Đăng xuất với khóa thật (tiêu chí 3).
