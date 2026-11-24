# Thăm dò: add-in gọi model Google qua Antigravity (23/11/2026)

## Kết luận

**Không làm.** Điều khoản hiện hành của Google Antigravity cấm dùng phần mềm bên thứ ba để truy cập dịch vụ (vi phạm thì có thể bị khoá tài khoản Antigravity), và hỏi đáp chính thức nói thẳng việc bọc `agy` trong ứng dụng khác không được hỗ trợ với gói cá nhân. Add-in AI Đa Nguồn là ứng dụng bên thứ ba. Về kỹ thuật cũng không có đường đi: Antigravity không mở địa chỉ HTTP nào cho trang web gọi, và Edge mặc định chặn trang `github.io` gọi vào máy cục bộ.

Đường hợp lệ đã có sẵn trong add-in: nguồn **Google Gemini** bằng API key của Google AI Studio (`providers.ts`, id `gemini`), chính là đường Google khuyên cho phần mềm bên thứ ba.

## Dự án API-CLI gọi Antigravity thế nào

Nguồn: `D:\API-CLI\BAN_GIAO.md` mục G8, `D:\API-CLI\tai-lieu\G8-0_NGHIEN_CUU_CHATGPT_GEMINI_WEB.md` mục 5b–5g.

- **Không qua địa chỉ, cổng hay HTTP.** Máy chủ Node cục bộ của API-CLI (`khung-chat/nha-cli.js`) chạy tiến trình `agy.exe` (Antigravity CLI 1.5.14, `%LOCALAPPDATA%\agy\bin`) ở chế độ không tương tác: `agy -p` hoặc đề bài qua stdin dạng `--input-format stream-json --output-format stream-json`.
- **Xác thực:** `agy` tự dùng thông tin đăng nhập Google đã lưu trong kho thông tin đăng nhập Windows (mục `gemini:antigravity`, dùng chung với bản desktop). API-CLI không cầm token nào.
- **Định dạng:** dòng JSON `{"event":"user","message":{"content":"…"}}` vào, dòng JSON sự kiện ra. Không phải Messages API của Anthropic, cũng không phải kiểu OpenAI.
- Mỗi lượt tốn khoảng 13.000–16.000 token vào (lời dặn và mô tả công cụ của `agy`), `agy` là tác tử có công cụ (đọc tệp, chạy lệnh); API-CLI phải chặn bằng luật deny và hồ sơ sạch mỗi lượt.
- Bản desktop còn có tiến trình `language_server.exe` nghe ở `127.0.0.1` (đo 23/11: cổng 60558, 60559, 51552) kèm mã CSRF trong `%USERPROFILE%\.gemini\antigravity\daemon\hub_*.json`; lệnh `agentapi` dùng đường này. Tài liệu chính thức chỉ dành `agentapi` cho "sidecar" do chính ứng dụng Antigravity khởi chạy. Tôi không đọc tệp `hub_*.json` (có mã CSRF).

## Bảng nghiệm thu

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| 1 | Tìm được dự án API-CLI và cách nó gọi Antigravity | Đạt | `D:\API-CLI` (nhánh G8, commit `dcf6be0` nghiệm thu thật 4/4): chạy tiến trình `agy`, không có điểm cuối HTTP |
| 2a | Mixed content: trang `https://…github.io` gọi `http://127.0.0.1` | Không cản | Lần đo đối chứng (tắt riêng lớp mạng nội bộ): máy chủ thử ở `http://127.0.0.1` và `http://localhost` nhận đủ `OPTIONS` + `POST`, trang đọc được 401. `bang-chung/do-doi-chung-tat-lna.json` |
| 2b | Private / Local Network Access (bước dò trước) | **Bị chặn** ở Edge mặc định | Edge 158.0.4370.33 `--headless=new`, trang `https://dunghoang1985.github.io/ai-da-nguon/taskpane.html`: cả 5 địa chỉ báo "Permission was denied for this request to access the `loopback` address space"; máy chủ thử **không nhận yêu cầu nào**, kể cả khi trả `Access-Control-Allow-Private-Network: true`. Đây là quyền người dùng phải cho, không phải tiêu đề máy chủ. `bang-chung/do-mac-dinh.json` |
| 2c | CORS ở cổng cục bộ của Antigravity | **Bị chặn** | Đối chứng: bước dò trước tới `127.0.0.1:60559` không có `Access-Control-Allow-Origin` |
| 2d | Trong khung WebView2 thật của Office | Chưa kiểm được | Không đo vì mục 3 đã chặn; WebView2 có cho hỏi quyền mạng nội bộ không thì tuỳ Office |
| 3 | Điều khoản cho phép ứng dụng bên thứ ba dùng quyền model của Antigravity | **Không** | Đọc ngày 23/11/2026 bằng Edge không cửa sổ (trang không ghi ngày cập nhật): `https://antigravity.google/terms` mục 6 nói dùng phần mềm, công cụ, dịch vụ bên thứ ba để truy cập dịch vụ là "a breach of this Agreement", có thể bị đình chỉ hoặc chấm dứt tài khoản Antigravity và Gemini CLI; cấm cả dùng dịch vụ cùng sản phẩm không do Google cung cấp. `https://antigravity.google/docs/faq/`: chế độ `agy -p` dành cho viết script trực tiếp và CI; bọc `agy` trong tác tử hoặc ứng dụng bên thứ ba "is not supported under consumer plans"; Google khuyên dùng API key Gemini Enterprise hoặc Google AI Studio |
| 4 | Khớp vào `providers.ts` / `dialect` / nguồn tự khai không thêm máy chủ riêng | Không áp dụng | Kể cả khi điều khoản cho phép: `agy` không có HTTP nên phải có máy chủ cục bộ chuyển tiếp (trái quy tắc "không thêm máy chủ riêng"); `dialect` chỉ mô tả khác biệt của nguồn HTTP kiểu Anthropic, không chạy được tiến trình |
| 5 | Không đưa key, token vào tệp | Đạt | Tệp đo chỉ gửi `Bearer khong-phai-key`; không đọc `KEY.txt` của API-CLI, không đọc `hub_*.json` |
| 6 | Không chiếm màn hình | Đạt | Edge `--headless=new`, hồ sơ tạm riêng, xoá sau khi đo; không bật Office, không chụp màn hình |

Cách đo (`bang-chung/do-cuc-bo.mjs`): máy chủ thử Node ở `127.0.0.1` với ba chế độ (CORS + cho phép mạng riêng · chỉ CORS · không tiêu đề), cộng cổng `language_server` của Antigravity chỉ để dò (không token). Mỗi địa chỉ một `POST` từ trang add-in thật, bộ tiêu đề như Đợt 6: `Content-Type: application/json`, `Authorization: Bearer …`, `anthropic-version`. Lần đối chứng thêm `--disable-features=LocalNetworkAccessChecks,PrivateNetworkAccessSendPreflights` để tách lớp chặn.

## Việc nên biết thêm (ngoài phạm vi add-in)

Câu hỏi đáp nêu trên cũng chạm tới chính dự án API-CLI: máy chủ của nó bọc `agy` để giao việc cho tác tử điều phối, mà Google ghi việc bọc `agy` trong ứng dụng khác là không được hỗ trợ với gói cá nhân. Đây là vùng chưa rõ (API-CLI là script cá nhân của anh, chạy trên máy anh, không phân phối), nhưng có rủi ro tài khoản Google bị khoá. Nên báo cho phiên của dự án API-CLI tự xem xét; tôi không sửa gì bên đó.

## Đề xuất để anh duyệt

**Câu 1. Add-in có làm đường Antigravity không?**
- [ ] Không làm, dùng nguồn Google Gemini bằng API key AI Studio đã có sẵn (Khuyến nghị)
- [ ] Hỏi lại Google (diễn đàn Antigravity) rồi tính sau
- [ ] Vẫn làm qua máy chủ cục bộ (trái điều khoản Google và trái quy tắc không thêm máy chủ riêng; không khuyến nghị)

**Câu 2. Rủi ro điều khoản ở dự án API-CLI (bọc `agy`)?**
- [ ] Báo phiên dự án API-CLI xem lại, cân nhắc tắt công tắc `tacTuAgy` (mặc định tắt, anh đang bật) (Khuyến nghị)
- [ ] Biết rồi, giữ nguyên

## Anh chọn (23/11/2026)

Câu 1: **không làm** đường Antigravity cho add-in; dùng nguồn Google Gemini bằng API key AI Studio đã có sẵn. Câu 2 (rủi ro điều khoản ở API-CLI): anh biết rủi ro, giữ nguyên.
