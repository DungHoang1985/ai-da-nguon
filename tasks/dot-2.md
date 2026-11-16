# Đợt 2 — Đăng nhập và chi phí (v4.1.0 → v4.2.0)

Lập 15/11/2026 (phiên Claude Code trên máy, nhánh `claude/dot-2-4268ba`, gốc `572f40f` = `origin/main`). Việc và tiêu chí gốc: mục "Đợt 2" và mục 4 "Đăng nhập thay cho dán API key" của `tasks/ke-hoach-v4.md`. Bằng chứng: `tasks/dot-2-bang-chung/`.

## Trạng thái

- [x] Đọc tài liệu dự án, đối chiếu tài liệu chính thức hiện hành (bảng dưới).
- [x] Mã, kiểm thử máy, thử trong Office tới ngay trước bước anh phải đăng nhập (16/11/2026).
- [x] Soát mã độc lập (0 cao, 4 trung bình, 8 thấp), sửa cả 12, chạy lại mọi bài trong Word.
- [ ] Việc cần anh ở máy (mục cuối): tiêu chí 1, 2, 3, 4 còn **Chưa kiểm được**.
- [ ] Hỏi anh, phát hành 4.2.0, kiểm bản đã phát hành, dọn dẹp.

Mã ở nhánh `claude/dot-2-4268ba` (chưa đẩy lên `main`). Bảng nghiệm thu ở mục "Kết quả" cuối tệp.

## Tiêu chí nghiệm thu (nêu trước khi làm)

Năm dòng đầu lấy đúng bảng của kế hoạch; dòng 6 là điều kiện chung của mọi đợt.

| # | Tiêu chí | Kiểm bằng |
|---|---|---|
| 1 | Từ trạng thái chưa có key OpenRouter: bấm một nút → đăng nhập → danh sách model hiện → sửa được một đoạn Word | Chạy thật |
| 2 | Lối dự phòng: chép mã từ trang, dán vào add-in, kết quả như trên | Chạy thật |
| 3 | Đăng xuất xóa sạch khóa trên máy | Chạy thật |
| 4 | Trang Thanh toán của Console hiện tín dụng; một yêu cầu chạy xong, số dư giảm | Ảnh trước và sau |
| 5 | Key không có số dư: thông báo tiếng Việt rõ lý do, hội thoại còn nguyên | Chạy thật |
| 6 | `npm test` đạt hết (đầu đợt 291), `npx tsc --noEmit -p .` sạch, `npm run build` đạt; Tự kiểm tra Word 34/34, Excel 22/22, PowerPoint 14/14 không giảm | Kiểm thử máy + Tự kiểm tra |

Tiêu chí 1, 2, 3, 4 cần chính anh đăng nhập hoặc mở trang Thanh toán: khi anh chưa ngồi trước máy thì ghi **Chưa kiểm được**, không ghi Đạt. Phần phiên tự kiểm được của các tiêu chí đó (hộp thoại mở đúng trang, đúng tham số; lối dán mã nhận và từ chối đúng; Đăng xuất xóa sạch trên kho thử) ghi thành dòng riêng với đúng mức bằng chứng.

Ba mức bằng chứng: **Tự kiểm tra** (nút trong add-in, chạy trong Office) · **Chạy thật** (giao việc cho model thật trong Office) · **Kiểm thử máy** (`npm test`, ngoài Office). Thử giao diện của bản đóng gói trong trình duyệt ngoài Office với nhà cung cấp giả lập **không phải** Chạy thật; ghi đúng tên là "Thử giao diện ngoài Office".

Quy tắc giữ nguyên: lịch sử hội thoại chỉ thêm, lời nhắc hệ thống cố định; không thêm máy chủ riêng; kho công khai.

## Quy tắc riêng của đợt này về tài khoản và khóa

- Phiên không gõ mật khẩu, không đăng nhập thay anh, không tạo tài khoản, không tạo hay thu hồi key, không nhận tín dụng, không bấm đồng ý điều khoản, không đụng tới thanh toán.
- Không đọc, không in kho lưu của khung add-in (có key thật). Không in key, mã đăng nhập, token ra màn hình, nhật ký hay tệp bằng chứng. Tệp bằng chứng không có email, tên tài khoản, số dư.
- Không xóa, không thay key thật đang lưu khi anh chưa đồng ý. Thử Đăng xuất và trạng thái "chưa có key" trên kho thử (bản đóng gói mở ở địa chỉ khác, kho riêng, chỉ có chuỗi giả).
- Mã bí mật của luồng đăng nhập (mã kiểm `code_verifier`) chỉ nằm trong bộ nhớ của khung, không ghi ra kho lưu, không nằm trong địa chỉ trang, không vào lịch sử hội thoại, nhật ký hay thông báo lỗi.

## Đối chiếu tài liệu chính thức (15/11/2026)

| Điều cần biết | Nguồn | Kết quả |
|---|---|---|
| Luồng đăng nhập OpenRouter | `openrouter.ai/docs/guides/overview/auth/oauth` | Mở `https://openrouter.ai/auth` với `callback_url`, `code_challenge`, `code_challenge_method=S256`, tùy chọn `key_label`, `state`. Sau khi người dùng đồng ý, trang chuyển về `callback_url?code=…`. Đổi mã: `POST https://openrouter.ai/api/v1/auth/keys` với `code`, `code_verifier`, `code_challenge_method` → `{ key }`. Mã dùng một lần, hết hạn sau 10 phút. Lỗi: 400 sai phương thức thách thức, 403 mã hoặc mã kiểm sai / hết hạn, 405 sai kiểu gọi |
| Lối dự phòng dán mã | cùng trang, mục ứng dụng không có địa chỉ nhận | Bỏ `callback_url` thì trang OpenRouter **hiện mã để chép**; bắt buộc có `code_challenge`; `state` không được trả về |
| Xem hạn mức của key | `openrouter.ai/docs/api/reference/limits` | `GET /api/v1/key`: `limit`, `limit_remaining` (trần riêng của key, `null` = không đặt trần), `usage`, `usage_daily`, `usage_monthly`, `is_free_tier`, `free_model_daily_requests { used, limit, remaining }`. Model miễn phí: 20 lượt/phút; 50 lượt/ngày nếu chưa từng nạp đủ 10 credit, 1.000 lượt/ngày nếu đã nạp. Số dư âm thì cả model miễn phí cũng trả 402 |
| Xem số dư tài khoản | `openrouter.ai/docs/api/api-reference/credits/get-credits` | `GET /api/v1/credits` nay **đòi khóa quản lý**; key thường bị 403. **Lệch với kế hoạch** (kế hoạch ghi "xem hạn mức gọi được từ trình duyệt" — đúng với `/key`, không đúng với số dư tài khoản). Cách làm: thẻ hạn mức gọi `/key` trước; thử `/credits`, được thì hiện số dư tài khoản, bị 403 thì hiện trần còn lại của key (nếu có đặt trần), mức đã dùng hôm nay / tháng này, số lượt miễn phí còn trong ngày, kèm đường dẫn mở trang số dư. Kiểm lại bằng lần gọi thật trong Office |
| Mã lỗi OpenRouter | `openrouter.ai/docs/api/reference/errors-and-debugging`, trang limits | 401 key sai, bị tắt hoặc phiên hết hạn; 402 hết credit (`error.metadata.limit_source`: `openrouter_credits`, `openrouter_key_limit`, `openrouter_in_flight_budget` — loại cuối là tạm thời, chờ rồi gửi lại); 403 ngân sách do quản trị đặt hoặc bị kiểm duyệt; 429 hết hạn mức; lỗi giữa luồng nằm trong sự kiện SSE |
| Tín dụng API kèm gói (bài 17154008) | `support.claude.com/en/articles/17154008`, `platform.claude.com/docs/en/about-claude/api-credits-for-subscribers` | **Ưu đãi còn**: Max 100 USD/tháng, Ultra 200 USD, Team gộp tối đa 500 USD; Free, Pro, Enterprise không có. Điều kiện: gói đang hoạt động và đã đủ 7 ngày. Nhận ở claude.ai → Settings → Billing → mục "API credits" → chọn tổ chức Console nhận tín dụng (liên kết là đồng ý điều khoản tín dụng bổ sung). **Mỗi gói chỉ liên kết một tổ chức, không tự đổi lại được.** Tín dụng hiện ở mục "Promotional credits" của `platform.claude.com/settings/billing`, hết hạn cuối mỗi kỳ, không cộng dồn; dùng cho Claude API, không dùng cho Claude Code. Hết tín dụng mà không có tín dụng mua thêm: yêu cầu dừng tới kỳ sau, không trừ vào gói |
| Thông báo hết tín dụng của Anthropic | trang tín dụng ở trên, `platform.claude.com/docs/en/api/errors` | Câu trả về: "Your credit balance is too low to access the Anthropic API. Please go to Plans & Billing to upgrade or purchase credits." Mã: 400 `invalid_request_error` (cũng dùng khi chạm trần chi tiêu tự đặt), 401 `authentication_error` (key sai, bị thu hồi, hết hạn), 402 `billing_error`, 403 `permission_error`, 429 gồm cả chạm trần chi tiêu của bậc. 401 có câu mở đầu "CORS requests are not allowed for this Organization" là do thiết lập của tổ chức, không phải sai key |
| Key Anthropic gắn workspace | `platform.claude.com/docs/en/manage-claude/authentication` | Tạo ở `platform.claude.com/settings/keys` → Create key: đặt tên, chọn hạn dùng, "Linked account" là chính mình, chọn một workspace. Key **không gắn workspace** thì mọi yêu cầu phải kèm tiêu đề `anthropic-workspace-id`, thiếu là 400 kèm câu "This API key is not scoped to a workspace…". Add-in không gửi tiêu đề đó → hướng dẫn anh chọn workspace khi tạo key, và add-in dịch đúng lỗi này |
| Hộp thoại của Office | `learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins` (cập nhật 28/09/2026) | `Office.context.ui.displayDialogAsync(url)`: trang đầu **phải cùng tên miền** với khung add-in và dùng HTTPS; sau đó được chuyển sang trang HTTPS khác. Trang trong hộp thoại gửi chuỗi về khung bằng `Office.context.ui.messageParent` (trang phải nạp office.js; cùng tên miền thì không cần `targetOrigin`). Mỗi khung chỉ mở được một hộp thoại. Không dùng `window.open`; mở trình duyệt ngoài bằng `Office.context.ui.openBrowserWindow`. Không cần sửa manifest (không cần `AppDomains` cho trang trong hộp thoại) → **không phải cài lại add-in** |

## Thiết kế

- **Khối đăng nhập dùng chung** `src/core/oauth.ts` (thuần, kiểm thử được): tạo mã kiểm và mã thách thức S256, dựng địa chỉ đăng nhập (có và không có `callback_url`), đọc mã từ địa chỉ trả về, nhận mã người dùng dán (chấp nhận cả khi dán nguyên địa chỉ), đổi mã lấy key, dịch lỗi. Nhà cung cấp khai bằng cấu hình (`oauth` trong `src/core/providers.ts`): thêm nguồn cùng kiểu sau này chỉ thêm cấu hình.
- **Trang hộp thoại** `dang-nhap.html` (cùng tên miền với add-in): lần mở đầu nhận mã thách thức rồi tự chuyển tới `openrouter.ai/auth` (tên miền đích cố định trong mã, không nhận địa chỉ đích từ tham số); khi OpenRouter chuyển về kèm `code` thì gửi về khung bằng `messageParent`. Nếu không gửi được (mở ngoài Office) thì hiện mã để chép.
- **Lối dự phòng**: nút mở trang OpenRouter trong trình duyệt ngoài, không có `callback_url` → OpenRouter hiện mã → anh chép, dán vào ô trong add-in. Cùng một cặp mã kiểm / thách thức với lối hộp thoại.
- **Thẻ OpenRouter trong Cài đặt** (`src/taskpane/components/SignInCard.tsx`): chưa có khóa thì hiện nút Đăng nhập, lối dán mã và ô dán key như cũ; có khóa thì hiện hạn mức, cảnh báo, nút Đăng xuất, nút Đăng nhập lại. Màn chào (chưa có key nào) có nút Đăng nhập OpenRouter ngay.
- **Đăng xuất**: xóa khóa OpenRouter và danh mục model đã lưu của nguồn đó khỏi kho của add-in. Khóa không tự thu hồi được ở phía OpenRouter (việc đó cần khóa quản lý): thẻ nói rõ và đưa đường dẫn tới trang quản lý khóa.
- **Hạn mức** `src/core/credits.ts` (thuần): đọc `/key` (và `/credits` nếu được), ra các dòng hiển thị và mức cảnh báo (còn ≤ 5 lượt miễn phí trong ngày; trần của key còn dưới 1 USD; hết). Khung chat hiện cảnh báo ngắn khi đang dùng OpenRouter và hạn mức thấp.
- **Lỗi key và tiền** (`src/core/errors.ts`): phân loại theo nhà cung cấp — hết tiền / hết tín dụng, key sai, key hết hạn hoặc bị thu hồi, hết lượt miễn phí trong ngày, chạm trần của key, key Anthropic chưa gắn workspace, tổ chức chặn gọi từ trình duyệt. Mỗi loại một câu tiếng Việt nêu lý do và việc nên làm (đổi nguồn ở nhãn model cạnh ô nhập, hoặc mở Cài đặt). Lỗi hết tiền không tự thử lại (trước đây Gemini trả 429 nên bị thử lại 3 lần vô ích). Hội thoại giữ nguyên: lượt lỗi không làm đổi lịch sử, đổi nguồn rồi gửi tiếp được.

## Các bước theo thứ tự, mỗi bước kiểm bằng gì

| Bước | Việc | Kiểm bằng |
|---|---|---|
| 1 | Viết tệp này, commit trước mọi commit mã | `git log` |
| 2 | `src/core/oauth.ts` + cấu hình OpenRouter | Kiểm thử máy (mã thách thức đúng véc-tơ chuẩn RFC 7636, địa chỉ đúng tham số, nhận và từ chối mã dán, đổi mã với phản hồi giả lập 200 / 400 / 403 / 405) |
| 3 | `src/core/credits.ts` | Kiểm thử máy (phản hồi mẫu `/key`: miễn phí còn nhiều, còn ít, hết; key có trần; `/credits` 403) |
| 4 | `src/core/errors.ts`: phân loại lỗi key và tiền của Anthropic, OpenRouter, Gemini, OpenAI, DeepSeek, Groq; lỗi hết tiền không thử lại | Kiểm thử máy (câu lỗi thật của từng nhà cung cấp lấy từ tài liệu và từ bằng chứng các đợt trước) |
| 5 | Trang hộp thoại `dang-nhap.html` + đóng gói | `npm run build` có `dang-nhap.html`; kiểm thử máy cho phần thuần |
| 6 | Giao diện: thẻ OpenRouter trong Cài đặt, nút ở màn chào, cảnh báo hạn mức và gợi ý đổi nguồn trong khung chat | Thử giao diện ngoài Office (kho thử, nhà cung cấp giả lập) |
| 7 | **Thử hộp thoại trong Office** (thử trước theo mục 8 của kế hoạch): hộp thoại mở được không, tới đúng `openrouter.ai/auth` với đúng tham số không, trang trả về gửi được mã về khung không (dùng mã giả → add-in báo đúng lỗi "mã không đúng", khóa đang lưu không đổi) | `kiem-thu/cdp.mjs` (bản cục bộ nạp vào khung) |
| 8 | Chạy thật ca hết tiền với Gemini (tài khoản đang báo hết tiền): thông báo tiếng Việt, hội thoại còn nguyên, đổi sang nguồn khác gửi tiếp được | Chạy thật (tiêu chí 5) |
| 9 | Thẻ hạn mức với key OpenRouter đang lưu: gọi `/key` thật trong Office, thẻ hiện đủ dòng (bằng chứng che số) | Chạy trong Office |
| 10 | Tự kiểm tra Word, Excel, PowerPoint trên bản cục bộ | Tự kiểm tra (tiêu chí 6) |
| 11 | Soát mã độc lập: tác tử khác đọc `git diff 572f40f..HEAD -- src test`, chú ý riêng key / mã đăng nhập có lọt vào nhật ký, lịch sử hội thoại, địa chỉ trang, thông báo lỗi, tệp chia sẻ hay tệp sao lưu không; sửa; chạy lại | Báo cáo trong thư mục bằng chứng |
| 12 | Cập nhật `tasks/HANDOFF.md`, `CLAUDE.md`, `docs/so-sanh-claude-chinh-thuc.md`; soát thư mục bằng chứng; **hỏi anh trước khi đẩy `main`** | — |
| 13 | Khi anh ngồi trước máy: các việc ở mục dưới (tiêu chí 1–4) | Chạy thật + ảnh |
| 14 | Sau khi anh duyệt: phát hành 4.2.0, chạy lại Tự kiểm tra trên bản đã phát hành, dọn dẹp (xóa hội thoại thử, trả nguồn về OpenRouter, đóng Office thử, đóng cổng gỡ lỗi) | Tự kiểm tra trên bản phát hành |

Bước không đạt: làm lại riêng bước đó tối đa 3 lần rồi dừng và báo.

## Điểm phải dừng và hỏi anh (ô tích chọn)

- Trước khi đẩy lên `main` (phát hành 4.2.0).
- Trước khi xóa hay thay key thật đang lưu.
- Trước mọi bước cần anh đăng nhập, tạo key, nhận tín dụng.
- Nếu hộp thoại Office không chạy được luồng đăng nhập và phải chọn cách khác lệch kế hoạch (lối dán mã vẫn là lối chắc chắn đã có trong kế hoạch).
- Trước khi đóng Office đang có tài liệu không phải tệp thử.

## Việc buộc phải có anh ở máy

Gom thành **một lần ngồi máy, khoảng 15 phút**. Phiên chuẩn bị sẵn Word với tệp thử và khung add-in; anh chỉ làm các cú bấm dưới đây, phiên đọc kết quả bằng mã. Phiên không nhìn và không ghi lại mật khẩu, key hay số dư của anh.

### A. Đăng nhập OpenRouter bằng một nút (tiêu chí 1)

Việc này thay khóa OpenRouter đang lưu trên máy bằng khóa mới do OpenRouter cấp qua đăng nhập. Khóa cũ vẫn còn hiệu lực ở phía OpenRouter (anh giữ hay xóa tùy ý ở trang quản lý khóa).

1. Trong khung add-in, bấm nút bánh răng (Cài đặt). Ở thẻ **OpenRouter**, bấm **Đăng xuất**, rồi bấm **Xóa khóa trên máy** để xác nhận. Đúng là: thẻ hiện "Đã xóa khóa OpenRouter khỏi máy này…" và lại có nút Đăng nhập; nhãn model cạnh ô nhập không còn model OpenRouter.
2. Bấm **Đăng nhập OpenRouter**. Một cửa sổ nhỏ của Office mở trang openrouter.ai.
3. Cửa sổ đó chưa biết anh là ai nên hiện trang **"Sign Up"** (đã kiểm 16/11/2026). Anh đã có tài khoản: kéo xuống cuối, bấm **Sign in**, rồi tự đăng nhập tài khoản OpenRouter của anh. Trang có ba nút GitHub, Google, MetaMask và ô email / mật khẩu.
4. Trang hỏi cho phép ứng dụng dùng tài khoản (nhãn khóa "AI Da Nguon (Office)"): bấm **Authorize**. Đúng là: cửa sổ tự đóng, thẻ OpenRouter hiện "Đã đăng nhập · n model" và các dòng hạn mức.
5. Bấm **Quay lại**. Phiên sẽ chọn một model OpenRouter gọi được công cụ, gửi một yêu cầu sửa một đoạn trong tệp thử và ghi kết quả.

Nếu ở bước 3 trang báo không đăng nhập được trong cửa sổ này (một số cách đăng nhập, ví dụ bằng tài khoản Google, có thể bị chặn trong cửa sổ nhúng): đóng cửa sổ và làm mục B. Đó là lý do có lối dự phòng.

### B. Lối dự phòng: chép mã, dán vào add-in (tiêu chí 2)

1. Ở thẻ OpenRouter bấm **Đăng xuất** → **Xóa khóa trên máy** (để về trạng thái chưa có khóa).
2. Bấm **Mở trang lấy mã trong trình duyệt**. Trình duyệt mặc định mở trang openrouter.ai; đăng nhập nếu được hỏi, rồi bấm **Authorize**.
3. Trang hiện một mã. Bấm nút chép (Copy) cạnh mã.
4. Quay lại Word, bấm vào ô **Dán mã vào đây** trong thẻ OpenRouter, nhấn Ctrl+V, rồi bấm **Xác nhận mã**. Đúng là: thẻ hiện "Đã đăng nhập · n model".
5. Phiên gửi một yêu cầu sửa đoạn như ở mục A.

Mã chỉ dùng một lần và hết hạn sau 10 phút; phải dán vào đúng khung add-in đã mở trang lấy mã (không tải lại khung giữa chừng).

### C. Đăng xuất (tiêu chí 3)

1. Bấm **Đăng xuất** → **Xóa khóa trên máy**. Đúng là: thẻ về "Chưa đăng nhập".
2. Phiên kiểm bằng mã rằng add-in không còn khóa OpenRouter (chỉ kiểm "còn hay không", không đọc giá trị) và danh mục model của OpenRouter đã bị xóa.
3. Bấm **Đăng nhập OpenRouter** lần nữa (như mục A) để dùng tiếp.

### D. Tín dụng API kèm gói Max và key Anthropic (tiêu chí 4)

Điều kiện theo tài liệu ngày 15/11/2026: gói Max đang hoạt động và đã đủ 7 ngày. **Lưu ý trước khi bấm: mỗi gói chỉ liên kết được một tổ chức Console và không tự đổi lại được** (muốn đổi phải nhờ bộ phận hỗ trợ) — chọn đúng tổ chức anh định dùng lâu dài.

1. Mở trình duyệt, vào `claude.ai`, đăng nhập. Bấm tên anh ở góc dưới bên trái → **Settings** → **Billing**.
2. Tìm mục **API credits**. Nếu không thấy mục này: gói chưa đủ điều kiện hoặc ưu đãi chưa tới tài khoản; dừng ở đây và báo phiên.
3. Chọn tổ chức Console sẽ nhận tín dụng (hoặc tạo mới ngay trong bước này), đọc điều khoản tín dụng, rồi bấm nút liên kết. Việc bấm liên kết là đồng ý điều khoản — anh tự quyết.
4. Mở `platform.claude.com/settings/billing`. Đúng là: mục **Promotional credits** hiện 100 USD kèm ngày hết hạn. **Chụp màn hình trang này** (ảnh "trước").
5. Mở `platform.claude.com/settings/keys` → **Create key**. Đặt tên (ví dụ "AI Da Nguon"), chọn hạn dùng, **Linked account** để là anh, và **chọn một workspace** (Default là được; key không gắn workspace thì add-in không gọi được). Bấm tạo, rồi bấm chép key.
6. Trong add-in: Cài đặt → thẻ **Anthropic (Claude)** → dán key vào ô → **Lưu**. Đúng là: "Đã kết nối · n model".
7. Phiên gửi một yêu cầu ngắn bằng model Claude.
8. Tải lại trang `platform.claude.com/settings/billing`. Đúng là: số dư tín dụng giảm (có thể chậm vài phút). **Chụp màn hình lần hai** (ảnh "sau").

Hai ảnh là bằng chứng của tiêu chí 4. Ảnh có số dư và có thể có tên tổ chức: **không đưa vào kho công khai**; phiên chỉ ghi "đã xem ảnh trước và sau, số dư giảm" vào bảng nghiệm thu, anh giữ ảnh trên máy.

### E. Không cần anh

Ca "key không có số dư" (tiêu chí 5) dùng tài khoản Gemini đang báo hết tiền; phiên tự chạy (đã chạy 16/11/2026, đạt).

### Lưu ý sau lần thử 16/11/2026

- Bài thử lối dán mã đã **bấm thật một lần** nút "Mở trang lấy mã trong trình duyệt": trên máy anh có một thẻ trình duyệt ở trang openrouter.ai do bài thử mở. Thẻ đó vô hại (chưa đăng nhập, chưa bấm gì); anh đóng đi.
- Ở bước "sửa một đoạn Word": model miễn phí anh đang chọn (Baidu ERNIE 5.5 Flash free) có lần trả lời "Đã sửa…" mà **không gọi công cụ nào**, văn bản không đổi (bằng chứng 09a). Cùng yêu cầu, model NVIDIA Nemotron 3 Ultra (free) sửa đúng, có Track Changes (bằng chứng 09b). Khi làm mục A và B phiên sẽ chọn model đã kiểm là gọi được công cụ, xong trả lại model của anh.

---

## Kết quả (16/11/2026, bản cục bộ 4.2.0 nạp vào khung, chưa phát hành)

Mã: 5 commit `c72805f` → `e352118` trên nhánh `claude/dot-2-4268ba`. Bằng chứng: `tasks/dot-2-bang-chung/` (đã soát: không có key, mã đăng nhập, email, tên tài khoản Office; số trong thẻ hạn mức thay bằng `#`; hai số tiền còn lại trong tệp 01 là số của phản hồi giả lập).

### Bảng nghiệm thu

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| 1 | Từ trạng thái chưa có key OpenRouter: bấm một nút → đăng nhập → danh sách model hiện → sửa được một đoạn Word (Chạy thật) | **Chưa kiểm được** — bước đăng nhập và bấm Authorize phải do anh làm; chưa xóa khóa thật đang lưu vì chưa hỏi anh | — |
| 1a | Phần phiên tự kiểm được: hộp thoại của Office mở trang của add-in rồi tới `openrouter.ai/auth` với đúng `callback_url`, mã thách thức 43 ký tự, `S256`, `state`, nhãn khóa | **Đạt** (chạy trong Word thật, dừng trước bước đăng nhập) | 03 mục B; 08b; ảnh 02 |
| 1b | Đường về: trang trả về gửi mã về khung, khung kiểm `state`, đóng hộp thoại, gọi đổi mã; mã giả bị máy chủ từ chối với câu tiếng Việt đúng; khóa đang lưu không đổi | **Đạt** (chạy trong Word thật, mã giả; cả ca state sai) | 03 mục C; 08c |
| 1c | Màn chào khi chưa có key nào có nút Đăng nhập OpenRouter | **Đạt** (Thử giao diện ngoài Office, kho thử trống) | 01 mục A |
| 1d | Với khóa OpenRouter đang có: danh sách model hiện, sửa được một đoạn Word có Track Changes | **Đạt** (Chạy thật, OpenRouter · NVIDIA Nemotron 3 Ultra free: 3 bước công cụ, "150%" → "120%", một cặp xóa/thêm). Với model ERNIE 5.5 Flash free: **không đạt** — model nói đã sửa nhưng không gọi công cụ (xem Vấn đề mở) | 09b; 09a |
| 2 | Lối dự phòng: chép mã từ trang, dán vào add-in, kết quả như trên (Chạy thật) | **Chưa kiểm được** — cần anh đăng nhập ở trình duyệt và chép mã thật | — |
| 2a | Phần phiên tự kiểm được: nút mở trang lấy mã mở được trình duyệt mặc định; địa chỉ không có `callback_url` và `state`; ô dán mã từ chối chuỗi sai dạng, địa chỉ không có mã, API key dán nhầm (0 lần gọi mạng); mã giả hợp dạng bị máy chủ từ chối với câu tiếng Việt đúng | **Đạt** (chạy trong Word thật với mã giả; địa chỉ kiểm ở bản đóng gói ngoài Office) | 03 mục D; 08d; 01 mục B, G |
| 2b | Trọn luồng dán mã → nhận khóa → danh sách model → thẻ hạn mức, với nhà cung cấp giả lập | **Đạt** (Thử giao diện ngoài Office; không phải Chạy thật) | 01 mục D |
| 3 | Đăng xuất xóa sạch khóa trên máy (Chạy thật) | **Chưa kiểm được** với khóa thật — xóa khóa thật phải hỏi anh trước | — |
| 3a | Phần phiên tự kiểm được: Đăng xuất có hỏi xác nhận; sau đó kho thử không còn khóa, không còn danh mục model của nguồn đó, tệp sao lưu mới không còn khóa; nguồn khác giữ nguyên | **Đạt** (Thử giao diện ngoài Office trên kho thử với key giả + Kiểm thử máy) | 01 mục F; `test/unit/dot2.test.ts` nhóm "đăng xuất" |
| 4 | Trang Thanh toán của Console hiện tín dụng; một yêu cầu chạy xong, số dư giảm (ảnh trước và sau) | **Chưa kiểm được** — nhận tín dụng, tạo key và chụp trang Thanh toán phải do anh làm. Đã đối chiếu tài liệu 15/11/2026: ưu đãi còn (Max 100 USD mỗi tháng), điều kiện và từng cú bấm ở mục D | Bảng đối chiếu ở đầu tệp |
| 4a | Phần mã: add-in nhận ra thông báo hết tín dụng của Anthropic (câu trong tài liệu chính thức) và gợi ý đổi nguồn; key chưa gắn workspace và tổ chức chặn gọi từ trình duyệt có câu riêng | **Đạt** (Kiểm thử máy; chưa có ca Chạy thật vì key Anthropic của anh còn tiền) | `test/unit/dot2.test.ts` nhóm "phân loại lỗi key và tiền" |
| 5 | Key không có số dư: thông báo tiếng Việt rõ lý do, hội thoại còn nguyên (Chạy thật) | **Đạt** — Gemini (tài khoản hết tiền) trả 402; add-in gọi đúng 1 lần, 2 giây, báo "Google Gemini: tài khoản đã hết tiền (hết tín dụng)…", có nút Đổi nguồn mở bảng chọn; đổi sang Claude Haiku 5.5 rồi hỏi lại thì model trả đúng mã thử của lượt trước, đọc 14.193 token từ bộ nhớ đệm. Chạy hai lần (trước và sau khi sửa theo soát mã) | 04; 08e; ảnh 04b |
| 6 | `npm test` đạt hết, `npx tsc --noEmit -p .` sạch, `npm run build` đạt; Tự kiểm tra Word 34/34, Excel 22/22, PowerPoint 14/14 không giảm | **Đạt** — `npm test` **363/363** (291 cũ + 72 mới), tsc sạch, đóng gói đạt (có `dang-nhap.html`). Word **34/34** (trên mã đã sửa theo soát mã); Excel **22/22**; PowerPoint 13/14 lần đầu trên bài mới rồi **14/14** (đúng lỗi cũ "Biểu đồ gốc thêm một bố cục", có từ bản 3.6). Ghi nhận: lần chạy Word đầu tiên ra 28/34 (xem Vấn đề mở), ba lần sau đều 34/34 | 05 (Word); 06 (Excel); 07a, 07b (PowerPoint) |
| — | Thẻ hạn mức với khóa thật của anh: đọc được số dư tài khoản, trần của key, mức đã dùng hôm nay / tháng này, số lượt miễn phí còn trong ngày | **Đạt** (chạy trong Word thật; số được che) | 03 mục A; 08a |

Tóm lại: **tiêu chí 5 và 6 đạt; tiêu chí 1, 2, 3, 4 chưa kiểm được** vì cần chính anh đăng nhập, bấm Authorize, đồng ý xóa khóa thật, nhận tín dụng. Mọi phần của bốn tiêu chí đó mà phiên tự kiểm được đều đã kiểm và đạt.

### Lệch so với kế hoạch và tài liệu, đã đo thật

| Điều | Kế hoạch / tài liệu ghi | Đo được 16/11/2026 | Xử lý |
|---|---|---|---|
| Hộp thoại Office chạy luồng đăng nhập | "phải thử trước" (mục 8 của kế hoạch) | **Chạy được**: mở trang cùng tên miền, chuyển sang openrouter.ai, quay về, `messageParent` tới khung | Làm đúng phương án chính, vẫn giữ lối dán mã |
| Đổi mã với mã sai | 403 "Invalid code or code_verifier" | **400 "Invalid code"** | Câu báo nhận cả hai |
| Xem số dư tài khoản OpenRouter | `/credits` cần khóa quản lý | Với khóa đang lưu của anh **xem được** | Thẻ hiện số dư khi xem được; không được thì hiện dòng "key này không được xem" kèm đường dẫn |
| Kiểm key OpenRouter | — | `/models` trả 200 cả khi key sai; `/key` trả 401 "User not found." | Kiểm bằng `/key` trước khi lưu; key sai không ghi đè key đang dùng |
| Gemini hết tiền trả trước | 429 | **402** kèm "Your prepayment credits are depleted" | Nhận ra ở cả hai mã, không tự thử lại |
| Trang OpenRouter trong hộp thoại | — | Hiện "Sign Up" (cửa sổ chưa đăng nhập), có nút "Sign in" ở cuối | Ghi vào hướng dẫn từng cú bấm |
| Mở trình duyệt ngoài | — | `Office.context.ui.openBrowserWindow` có, không ghi đè được | Bài thử mặc định không bấm nút này |

### Soát mã độc lập

Tác tử khác (không tham gia viết mã) đọc `git diff 572f40f..HEAD -- src test webpack.config.js`, tự viết ca thử: **0 cao · 4 trung bình · 8 thấp**. Báo cáo nguyên văn: `tasks/dot-2-bang-chung/soat-ma-doc-lap.md`. Kết luận về bí mật: luồng đăng nhập không làm lộ khóa, mã đăng nhập hay mã kiểm; chỗ rò duy nhất là câu lỗi của nhà cung cấp có thể nhắc lại key (có từ trước đợt này). **Đã sửa cả 12** (commit `e352118`), thêm 14 kiểm thử máy, rồi chạy lại mọi bài trong Word:

| Mã | Phát hiện | Đã sửa thế nào |
|---|---|---|
| T1, Th1 | Câu lỗi của nhà cung cấp chèn thẳng vào thông báo, chi tiết lỗi và hội thoại lưu, không che key | `maskSecrets` trong `errors.ts` che mọi dạng key đã biết; áp ở `explainApiError`, lỗi giữa luồng, và ở `describeError` làm lớp chặn cuối; `redact` của luồng đăng nhập dùng chung |
| T2 | Key của nhà cung cấp khác dán nhầm vào ô mã bị gửi sang OpenRouter làm "mã" | Chuỗi có dạng key bị từ chối tại chỗ ("là API key, không phải mã đăng nhập"), không gọi mạng (đã đo trong Word: 0 lần gọi) |
| T3 | 402 "requires more credits, or fewer max_tokens… can only afford N" bị báo là hết tiền | Lý do riêng `low_credit`: "số dư còn ít, không đủ cho độ dài trả lời tối đa… (tài khoản chưa hết tiền)" |
| T4 | Thẻ hạn mức: kết quả của key cũ đè key mới; lỗi mạng của `/credits` bị nhớ thành "không được xem"; đăng nhập lại không tự đọc lại | Bộ nhớ đệm theo "thế hệ" key; chỉ kết luận "không được xem" khi máy chủ trả 401/403; kết quả bị bỏ thì thẻ tự đọc lại |
| Th2 | 429 kèm lời mời nạp tiền bị gán "hết tiền", mất tự thử lại | Với 429 chỉ nhận hết tiền khi có tín hiệu chắc chắn (mã lỗi, "depleted", "insufficient") |
| Th3 | 400 / 403 có chữ "api key" đâu đó bị đổ cho key | Chỉ nhận khi câu nói rõ về chính key |
| Th4 | Lỗi hết lượt theo phút kèm câu mời bị nhận là hết lượt theo ngày | Chỉ khớp `free-models-per-day` |
| Th5 | Bấm đăng nhập lúc đang đổi mã gây lỗi "state không khớp" | Khóa nút khi đang bận; mã về thì đóng hộp thoại |
| Th6 | Đăng xuất luôn báo "đã xóa"; lượt đổi mã đang chờ vẫn lưu khóa sau khi đăng xuất | Đọc lại kho rồi mới báo; đăng xuất hủy lượt đang chờ |
| Th7 | Key có trần nhưng thiếu số còn lại bị coi là đã chạm trần | Không đoán; chỉ báo hết khi máy chủ trả số ≤ 0 |
| Th8 | Đăng nhập lại không nhắc khóa cũ còn hiệu lực; nút mở trang không bắt lỗi | Thêm lời nhắc kèm đường dẫn trang quản lý key; bọc bắt lỗi |

Lỗi do chính các bài thử trong Word tìm ra (trước lần soát mã), đã sửa: sau khi bấm "Mở trang lấy mã…", nút "Xác nhận mã" bị khóa (dùng nhầm trạng thái "đang bận" cho dòng hướng dẫn).

Nghi vấn của lần soát chưa kiểm: office.js có ghi địa chỉ có `?code=` vào kho phiên của hộp thoại trước khi mã được gỡ khỏi thanh địa chỉ hay không (rủi ro thấp: mã dùng một lần, vô dụng khi thiếu mã kiểm nằm trong khung; hộp thoại đóng ngay sau khi gửi mã).

### Giới hạn đã biết và vấn đề mở

- **Model miễn phí nói "đã sửa" mà không gọi công cụ** (ERNIE 5.5 Flash free, bằng chứng 09a): add-in hiện chưa cảnh báo khi câu trả lời khẳng định đã sửa mà lượt đó không có thao tác ghi nào. Không thuộc phạm vi đợt này; đề xuất thêm cảnh báo ở đợt sau.
- **Tự kiểm tra Word lần đầu 28/34**: 5 bước lỗi `AccessDenied` (chèn nhận xét) và `GeneralException` (chèn ô điền mẫu), ngay sau bài Chạy thật với Gemini và sau khi trình duyệt ngoài vừa được mở; chạy lại ngay ra 34/34, và hai lần sau cũng 34/34. Chưa rõ nguyên nhân (nghi Word chưa sẵn sàng nhận lệnh chèn lúc cửa sổ vừa mất tiêu điểm). Các công cụ đó không bị sửa ở đợt này.
- Add-in không thu hồi được khóa ở phía OpenRouter (cần khóa quản lý): Đăng xuất chỉ xóa trên máy và chỉ đường tới trang quản lý key. Mỗi lần đăng nhập lại, OpenRouter cấp thêm một khóa mới.
- Đăng nhập bằng Google trong hộp thoại nhúng có thể bị Google chặn: chưa kiểm được (cần anh); có lối dán mã.
- Số dư tài khoản Anthropic không xem được bằng API (không có địa chỉ xem số dư): chỉ báo khi đã hết.
- Ô "Dán mã vào đây" là ô chữ thường (không che ký tự): mã dùng một lần và vô dụng khi thiếu mã kiểm; chuỗi có dạng API key bị từ chối.
- Tệp sao lưu tạo **trước** khi đăng xuất vẫn chứa khóa cũ (đúng thiết kế của Sao lưu); khôi phục kiểu gộp sẽ đưa khóa đó trở lại.
- Chưa có kiểm thử máy cho phần giao diện (`useSignIn`, `useCredits`, `dang-nhap.ts`): các phần này được kiểm bằng bài chạy trong Word và ngoài Office.

### Dọn dẹp sau khi thử (16/11/2026)

Đã xóa các hội thoại thử trong Word (lịch sử trở lại 3 hội thoại như trước khi thử), trả nguồn về OpenRouter và model của anh (ERNIE 5.5 Flash free), đóng Word / Excel / PowerPoint thử không lưu, cổng gỡ lỗi 9229 đã đóng, kho thử `127.0.0.1:8765` đã xóa. Còn lại trên máy: một thẻ trình duyệt ở trang openrouter.ai do bài thử mở (anh đóng đi); model đang nhớ của nguồn Anthropic là Claude Haiku 5.5 (bài thử chọn, như các đợt trước).
