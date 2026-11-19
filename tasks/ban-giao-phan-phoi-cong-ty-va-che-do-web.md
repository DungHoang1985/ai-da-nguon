# Bàn giao – Phân phối cho công ty, chi phí AI và chế độ web chat

Phiên thảo luận: 18/11/2026. Dự án liên quan: add-in "AI Đa Nguồn" (`D:\ADD IN`, bản 4.8.0) và khung chat API-CLI (`D:\API-CLI`).

**Trạng thái: chưa có kế hoạch nào được duyệt. Không có tệp mã nào bị sửa, không phát hành gì.** Phiên này chỉ đọc mã, đọc tài liệu và tra cứu. Tệp này gom toàn bộ nội dung để chat sau chọn hướng và xử lý.

Cách đọc nhãn: **[Đã đọc mã]** = đối chiếu trực tiếp tệp trong kho ngày 18/11/2026. **[Nguồn ngoài]** = trang web tìm được, chưa đối chiếu tài liệu chính thức trừ khi ghi khác. **[Chưa kiểm]** = chưa xác nhận.

---

## 1. Anh muốn gì (theo thứ tự đã nêu)

1. Đóng gói API-CLI cho người khác tự điền key của họ.
2. Dùng khung chat khi hạn mức gói Max hết.
3. Add-in cài offline có CLI phía sau; nhúng CLI vào add-in; ưu nhược của việc nhúng.
4. Có hỗ trợ trong Word trực quan như Claude Desktop; bộ skill nạp sẵn đã đủ chưa; nhúng "kho thư viện skill" dùng khi cần.
5. Chi phí là vấn đề, nhất là token free; đẩy lên GitHub thì nhiều người truy cập có sập không.
6. Cấp add-in cho toàn công ty, anh trả tín dụng, chỉ cho tài khoản được cấp hoặc máy được cài dùng; GitHub thì mở.
7. Dùng chung key cá nhân của anh có rủi ro gì.
8. Gemini free hay bị dừng sớm: do quá tải hay hết hạn mức.
9. Tận dụng chat web miễn phí thay cho API (anh từng dùng "claw", con cua); chế độ web không cần đăng nhập; chế độ web nhúng trong add-in, tự chuyển câu hỏi sang chat và lấy câu trả lời về thao tác trong Word.

Ba hướng đã nêu cho việc đóng gói add-in (anh **chưa chọn**): **A** đóng gói add-in như hiện có (không CLI); **A+** thêm kết nối MCP từ xa; **B** nhúng chế độ CLI.

---

## 2. Hiện trạng kỹ thuật đã xác nhận

**Add-in [Đã đọc mã]**
- Bản 4.8.0, chạy trong khung Office, phát hành từ GitHub Pages (kho công khai `DungHoang1985/ai-da-nguon`, nhánh `main` và `gh-pages`). Trí nhớ cũ ghi "Cloudflare Pages + Worker" là sai.
- Gọi nhà cung cấp trực tiếp từ khung add-in; key lưu cục bộ. "Nguồn tự khai" cho phép nhập địa chỉ https + key, kiểu gọi Anthropic (`dialect`).
- Đã có sẵn nguồn: Google Gemini, Groq, OpenRouter (kể cả kênh Claude), DeepSeek (kiểu Anthropic), MiniMax, Kimi, Anthropic, Cerebras… (`src/core/providers.ts`).
- Skill: 18 skill dựng sẵn (11 cái cho Word), không ghi số hiệu văn bản pháp luật (có kiểm thử chặn). Nhập skill bằng .zip (một hoặc nhiều skill, hoặc gói plugin Claude), tệp `references/` mỗi tệp tối đa 60.000 byte (bị cắt nếu dài hơn), danh mục tối đa 40 skill × 300 ký tự (`src/core/skills.ts`).
- Quy tắc trong `CLAUDE.md`: kho công khai nên không chứa bí mật; không có máy chủ riêng; các quy tắc Office; không xóa thư mục `Wef`. (Mới đọc 120 dòng đầu; đọc lại nguyên văn trước khi đổi.)

**Đường gọi Gemini [Đã đọc mã]**
- Add-in **không gửi `maxOutputTokens` cho Gemini**. Trần `maxOutput` chỉ dùng cho Anthropic (`llm.ts` dòng 165). Với Gemini chỉ gửi `thinkingConfig` khi bật công tắc "Suy nghĩ kỹ": `thinkingLevel: "high"` cho Gemini 3.x, ngân sách 24.576 hoặc 32.768 cho 2.5.
- Phân loại lỗi (`errors.ts`): 429 hết hạn mức **theo ngày** (câu lỗi có "PerDay") thì không tự thử lại; 429 theo phút, 408, 5xx và lỗi mạng thì tự thử lại. `finishReason = MAX_TOKENS` được coi là bị cắt; trả lời rỗng thì khung hiện kèm lý do nhà cung cấp báo. Hạn chờ im lặng 150 giây.
- **Không có tệp log gọi API.** Chỉ có thông báo lỗi trên khung, nút "Chép chi tiết lỗi", và các tệp JSON kết quả bài thử trong `tasks/dot-*-bang-chung/`. Sheet "Nhật ký AI" chỉ ghi thao tác sửa ô Excel.

**Ghi chép các bài thử Gemini 3.8 Flash [Đã đọc mã]**
- 17/11 (bản 4.4.1): sáng trả lời đúng một bài, chiều bị 429 (`tasks/dot-5.md`, B1).
- PDF 3 trang: 429. PDF 7 trang: máy chủ im lặng 150 giây (`tasks/dot-5.md`, B2).
- 18/11: bài PowerPoint 2 slide im lặng 150 giây ngay lần gọi đầu (`dot-7-bang-chung/ppt-gemini-2-slide.json`). Bài điền mẫu: lần gọi 1 chạy được (vào 12.274, ra 310 token), lần 2 im lặng (`dot-8-bang-chung/chay-that-gemini.json`).
- `HANDOFF.md` ghi chuỗi "503, im lặng, rồi 429 cho cả ngày", mỗi lượt Word khoảng 11–35 nghìn token.
- **Kết luận:** các bài thử cho thấy dừng sớm do giới hạn/quá tải của gói miễn phí, **không phải trần token** (không có ghi chép MAX_TOKENS nào ở Gemini). **Chưa phân biệt được "quá tải" với "bị giới hạn"** ở trường hợp im lặng vì không có mã HTTP. Chưa dò 80 bài lẻ trong `kiem-thu/cdp-bai`.

**API-CLI [Đã đọc]** (đọc `BAN_GIAO.md`, `KE_HOACH.md`): bọc `claude -p` với key Anthropic Console, sổ chi phí, trần theo ngày và theo dự án, bộ nhớ đệm 5 phút/1 giờ. Tín dụng API tặng dùng được với `claude -p` bằng key, không dùng được cho phiên Desktop. Quy tắc S7: không chia sẻ key, không mượn key người khác. Không phải lõi nhiều nhà cung cấp.

---

## 3. Kết luận theo chủ đề

### 3.1 Cấp add-in cho cả công ty

- **GitHub không phải nút thắt tải** [Nguồn ngoài, tìm trước đó]: Pages có mức mềm khoảng 1 GB dung lượng, 100 GB băng thông/tháng, 10 lần build/giờ; Pages miễn phí cần kho công khai. Nút thắt thật là **số lần gọi AI và giới hạn tốc độ của key dùng chung**.
- **GitHub mở nên không giới hạn truy cập được.** Phát hành tập trung qua M365 chỉ kiểm soát việc ai thấy add-in, không kiểm soát ai dùng được AI. Kiểm soát thật nằm ở **key**.
- **Tổ chức Console** [Nguồn ngoài, tài liệu Anthropic]: mời thành viên bằng email, vai trò `user`/`developer`/`billing`/`claude_code_user` (vai trò `admin` chỉ cấp trong Console). **Workspace**: tối đa 100 mỗi tổ chức, chỉ admin tạo; key gắn với workspace, không chuyển được; đặt hạn mức chi tiêu tháng và giới hạn tốc độ theo từng workspace; hạn mức của tổ chức vẫn áp dụng. Hạn mức tháng theo từng người chỉ có ở workspace Claude Code, nơi key tự cấp khi đăng nhập và không tạo tay được, nên **không dùng được cho add-in**. Hạn mức theo từng key: nguồn không chính thức nói chưa có, **[Chưa kiểm]** trên tài liệu chính thức.
- **Ba phương án:**
  - **P1. Console công ty + mỗi người một workspace** (khoảng 100 người). Mỗi người một key, một hạn mức tháng; thu hồi riêng từng người. Không cần đổi quy tắc. Nhược: key vẫn nằm trên máy người dùng (lộ thì thiệt hại tối đa bằng hạn mức của người đó).
  - **P2. Cổng trung gian** (LiteLLM, Portkey, Kong…): key thật giữ ở máy chủ, cấp key ảo theo người, hạn mức tháng, thu hồi, có thể giới hạn model. Cần một máy chủ công ty vận hành, nên **phải đổi quy tắc "không máy chủ riêng"**. **[Chưa kiểm]** cổng có nhận được kiểu gọi Anthropic của "nguồn tự khai" không.
  - **P3. Chỉ phát hành tập trung qua M365**: kiểm soát hiển thị, không kiểm soát dùng.

### 3.2 Dùng chung key cá nhân của anh: nên tránh

- Key nằm trên máy mọi nhân viên, lấy ra được; không thu hồi riêng từng người (đổi key = cả công ty ngừng).
- Không chia hạn mức theo người; một người chạy hàng loạt có thể dùng hết tiền; cả công ty chung một giới hạn tốc độ.
- Anh chịu trách nhiệm mọi hoạt động trên tài khoản (Điều khoản Thương mại, mục D.5). Một người vi phạm chính sách có thể khiến tài khoản bị tạm ngưng (mục I.3.a), anh mất luôn khung chat và API-CLI.
- Điều khoản: Điều khoản Thương mại cho dùng dịch vụ để chạy sản phẩm cho "khách hàng và người dùng cuối" của mình (A.1) nhưng cấm bán lại nếu chưa được duyệt (D.4(a)). Một bản Điều khoản cá nhân (bản lưu trữ) cấm chia sẻ khóa API. **[Chưa kiểm]** tài khoản của anh thuộc điều khoản nào; **[Chưa kiểm]** điều kiện của tín dụng khuyến mãi gắn với tài khoản anh.
- Dữ liệu pháp lý của nhân viên đi qua tài khoản cá nhân anh, không có nhật ký theo người.

### 3.3 Gói miễn phí: ai hào phóng nhất [Nguồn ngoài, các trang tổng hợp 8–11/2026, chưa đối chiếu trang chính thức]

| Nhà cung cấp | Mức miễn phí | Lưu ý |
|---|---|---|
| Google Gemini (AI Studio) | Flash, Flash-Lite, Gemma; không cần thẻ | Chất lượng cao nhất nhóm free, ngữ cảnh 1 triệu token. Google không còn công bố hạn mức từng model, xem trong AI Studio. Bậc free dùng nội dung để cải thiện sản phẩm Google. Hạn mức tính theo dự án, không theo key. |
| Groq | 30 yêu cầu/phút, 1.000/ngày, 200.000 token/ngày (gpt-oss-120b) | Nhanh nhất, hạn mức rõ; ngân sách token ngày nhỏ. |
| OpenRouter (model :free) | 20 yêu cầu/phút, 50/ngày; 1.000/ngày sau khi nạp 10 USD một lần | Thêm key hay tài khoản không tăng hạn mức; danh sách model free đổi liên tục (14 model ngày 2/11). |
| Mistral | Nguồn mâu thuẫn | Một nguồn: gói Experiment khoảng 1 tỷ token/tháng nhưng phải đồng ý cho dùng dữ liệu huấn luyện. Nguồn khác: từ khoảng 5/11 gói Free không còn kèm API. Cần kiểm trực tiếp. |
| Cerebras | Hết bậc free thường trực từ 16/7/2026 | Còn gói dùng thử. |
| Cloudflare Workers AI | 10.000 neurons/ngày | |
| Gemini CLI đăng nhập Google | Nhiều nguồn: đã ngừng 18/6/2026, chuyển sang Antigravity CLI | Trang cũ còn ghi 1.000 yêu cầu/ngày (lỗi thời). Khóa API free: 250 yêu cầu/ngày, 10/phút, chỉ Flash. |

Kết luận: không gói free nào đủ cho cả công ty; không nên đưa nội dung mật vào gói free (Gemini dùng dữ liệu; Mistral Experiment đòi đồng ý huấn luyện).

### 3.4 Gemini dừng sớm: cách phân biệt (đã giải thích với anh)

- Mã 429 (RESOURCE_EXHAUSTED) = hạn mức. Mã 503 (UNAVAILABLE) = dịch vụ tạm thời không sẵn sàng. Cách xử lý chung của Google: thử lại với chờ tăng dần.
- Trả về 200 nhưng cụt hoặc rỗng, `finishReason = MAX_TOKENS` = trần chung của suy luận cộng câu trả lời bị chạm. Với add-in, khả năng này **thấp** vì add-in không đặt trần cho Gemini; nhưng suy luận mặc định bật ở Gemini 3.x.
- Trường hợp "im lặng 150 giây": không có mã HTTP nên chưa phân biệt được. Cần bài đo ở mục 6.

### 3.5 Chat web thay cho API

**"Claw"** (em đoán là OpenClaw bản "zero-token"; anh nói "con cua", cần xác nhận): điều khiển trình duyệt (Playwright/CDP), lấy cookie và token phiên rồi gọi giao diện web của Claude, ChatGPT, Gemini…; lưu `auth.json` trên máy.
- **Không nên dùng:** điều khoản Claude cá nhân cấm bot/script trừ khi qua API key hoặc được cho phép rõ; Google có thể đình chỉ tài khoản (nguồn bên thứ ba, **[Chưa kiểm]** điều khoản gốc của Gemini Apps); lưu cookie phiên trên máy; cần máy chủ cục bộ (trái quy tắc); dự án tự nói chỉ hợp thử nghiệm cá nhân.

**Chat web không cần đăng nhập:**
- ChatGPT: có ở các nước được hỗ trợ; một hội thoại mỗi lần, không lưu lịch sử, không tải tệp, model cũ hơn.
- Gemini: có, chỉ trên web máy tính; không tải tệp, không lịch sử, không đổi model.
- Claude: không (luôn cần tài khoản). Perplexity: tìm kiếm cơ bản không cần tài khoản. Duck.ai: không cần tài khoản, chat được ẩn danh. Copilot: nguồn mâu thuẫn (có nguồn nói từ cuối 8/2026 bắt buộc đăng nhập).
- Không đăng nhập vẫn là gửi nội dung lên dịch vụ: không đưa hợp đồng, NDA thật.

**Chế độ web trong add-in: giới hạn kỹ thuật [Nguồn ngoài, tài liệu Microsoft/MDN, chưa thử trên máy anh]:**
- Nhúng trang chat bằng iframe: gần như chắc chắn bị chặn (X-Frame-Options / CSP `frame-ancestors`). Chưa đo từng trang.
- Hộp thoại Office (`displayDialogAsync`): trang đầu phải cùng tên miền với add-in nhưng có thể chuyển hướng sang tên miền khác; cách này hiển thị được chat cạnh Word và để người dùng tự đăng nhập. **[Chưa kiểm]**: Google thường chặn đăng nhập trong trình duyệt nhúng.
- Add-in **không đọc được nội dung trang khác** (khác nguồn gốc); `messageParent` chỉ nên gọi từ trang cùng tên miền với add-in. Nên "tự lấy câu trả lời về" cần thành phần ngoài (tiện ích trình duyệt hoặc máy chủ cục bộ).
- Clipboard trong khung add-in có thể bị chặn bởi chính sách quyền (có báo cáo ở Office trên web). Cần ô dán dự phòng.
- **Web chat không điều khiển trực tiếp được add-in** (chỉ trả chữ). Điều khiển gián tiếp được: add-in ghi định dạng trả lời cố định vào lời nhắc (ví dụ khối `<tool_call>` hoặc danh sách sửa đổi theo số đoạn), rồi đọc và thực hiện sau khi người dùng duyệt. Claw cũng làm kiểu này. Chỉ chạy một lượt; không có vòng lặp nhiều bước tự phản hồi.
- **Phương án "chuyển tay" (add-in soạn lời nhắc, anh dán sang chat, dán câu trả lời về, add-in xem trước rồi áp dụng):** khả thi, hợp lệ, không cần máy chủ, **nhưng anh đã loại vì quá thủ công** (hai lần dán mỗi lượt).

---

## 4. Quyết định đã có và lý do

| Quyết định | Lý do |
|---|---|
| Chưa chọn hướng A / A+ / B; chưa duyệt kế hoạch nào | Anh vẫn đang hỏi làm rõ; việc lớn, khó làm lại thì chờ duyệt. |
| Không dùng chung key cá nhân cho công ty | Rủi ro mục 3.2; theo quy tắc S7 của chính anh. |
| Không dùng "claw"/zero-token ở dạng hiện có | Rủi ro điều khoản và tài khoản, cần máy chủ cục bộ, lưu cookie phiên (mục 3.5). |
| Loại phương án "chuyển tay" | Anh đánh giá quá thủ công. Anh muốn phương án tự động tối đa. |
| Nếu cần đổi quy tắc để hoàn thành dự án thì **được xem xét** | Anh cho phép (18/11/2026); xem mục 5. |
| Gemini free chỉ để thử, không dùng làm thước đo chất lượng | Đã ghi ở `HANDOFF.md`; hết hạn mức thì ghi "Chưa kiểm được". |

---

## 5. Quy tắc có thể cần đổi để hoàn thành dự án

| Quy tắc | Vướng với | Nếu đổi | Rủi ro / lưu ý |
|---|---|---|---|
| Không có máy chủ riêng (`CLAUDE.md` add-in) | Cổng key theo người (P2); chế độ web tự động; nhúng CLI (hướng B) | Cho phép **một** dịch vụ nhỏ do công ty vận hành (cổng key). Với chế độ web tự động thì cần máy chủ cục bộ hoặc tiện ích trình duyệt. | Phải có người vận hành, bảo mật, chi phí. Chế độ web tự động còn kèm rủi ro điều khoản/tài khoản, nên tách riêng quyết định này. |
| Kho công khai không chứa bí mật | Kho skill pháp lý nội bộ, cấu hình riêng của công ty | Giữ kho công khai; nội dung nội bộ đóng gói .zip riêng tư, nhập cục bộ (tệp `references/` chia dưới 60 KB). Kho riêng thì mất Pages miễn phí. | Tách rõ phần công khai và phần nội bộ trước khi nhập. |
| S7: không chia sẻ key / mượn key người khác (API-CLI) | Dùng chung key cho công ty | Không đổi quy tắc; dùng tổ chức Console công ty (P1/P2). | Cần xác nhận điều khoản áp dụng cho tài khoản công ty. |
| "Giới hạn truy cập theo tài khoản hoặc máy" | GitHub mở, M365 chỉ kiểm hiển thị | Chuyển điểm kiểm soát sang key (workspace hoặc cổng). | Không đạt bằng cách giấu mã. |

---

## 6. Vấn đề mở

1. **Quy mô và quyền:** công ty bao nhiêu người dùng AI; anh có quyền quản trị M365 và Console của công ty không (anh chưa trả lời; quyết định chọn P1 hay P2).
2. **Chọn hướng đóng gói add-in** cho người khác: A, A+ hay B.
3. **Chế độ web tự động:** làm hay bỏ. Nếu làm, phải chấp nhận thành phần ngoài (máy chủ cục bộ/tiện ích) và rủi ro điều khoản; hoặc chấp nhận bản bán tự động. Chưa thử trên máy anh: iframe, hộp thoại Office, clipboard, đăng nhập Google trong trình duyệt nhúng.
4. **Chẩn đoán Gemini:** chưa có mã HTTP, `finishReason`, `usageMetadata` cho các lần im lặng.
5. **Thư viện skill pháp lý cho Word:** chưa làm; chưa đo số lần gọi AI/token có và không có skill; chưa thử nhập .zip có nhiều `references/` trong Word thật.
6. **Chưa kiểm:** điều khoản áp dụng cho tài khoản Console cá nhân của anh; điều kiện tín dụng khuyến mãi; tín dụng có dùng chung giữa các workspace; hạn mức theo từng key trong tài liệu chính thức; cổng trung gian nhận được kiểu gọi của "nguồn tự khai"; điều khoản gốc Gemini Apps; số liệu gói free (các trang tổng hợp, đổi nhanh).
7. **Chi phí:** chưa đo chất lượng tiếng Việt pháp lý của các nguồn trả phí giá thấp đã có sẵn trong add-in (DeepSeek, MiniMax, Kimi) để thay gói free. Đây là gợi ý, chưa bàn với anh.

---

## 7. Bước tiếp theo đề xuất

| # | Việc | Tiêu chí nghiệm thu | Model · effort |
|---|---|---|---|
| 1 | Anh trả lời mục 6.1 và chọn P1 hay P2 | Có quyết định bằng chữ | (anh) |
| 2 | Bài đo chẩn đoán Gemini: thêm ghi mã HTTP, `finishReason`, `usageMetadata`, thời gian chờ cho mỗi lần gọi Gemini, **chỉ ở chế độ thử, không ghi key**; chạy lại cùng bài bằng một key trả phí vài USD | Kiểm thử máy xanh; log có đủ trường; kết luận bằng bảng: hết im lặng khi trả phí = hạn mức free; vẫn im lặng = Google quá tải | Opus 5.5 · Low, chạy trong Claude Code mở ở `D:\ADD IN` |
| 3 | Thử P1 thật: tạo một workspace hạn mức nhỏ + một key, dùng qua "nguồn tự khai" | Chạy được; vượt hạn mức thì bị chặn; thu hồi key thì dừng | Opus 5.5 · Low |
| 4 | Nếu chọn P2: thiết kế cổng (key ảo, hạn mức, giới hạn model, nhật ký) và sửa quy tắc "không máy chủ riêng" | Kế hoạch + tiêu chí được anh duyệt trước khi làm | Opus 5.5 · Medium |
| 5 | Nếu muốn chế độ web tự động: bài đo kỹ thuật mục 6.3, rồi quyết định có chấp nhận rủi ro không | Bảng đo: iframe, hộp thoại, clipboard, đăng nhập | Opus 5.5 · Medium |
| 6 | Thư viện skill pháp lý (.zip riêng tư) | Đo số lần gọi AI/token có và không skill; tách phần công khai và nội bộ | Opus 5.5 · Medium |

Câu mở đầu cho chat mới: *"Đọc `D:\ADD IN\tasks\ban-giao-phan-phoi-cong-ty-va-che-do-web.md`, rồi cho tôi chọn việc tiếp theo ở mục 7."*

Khi các việc trên có kết quả dùng lại được (quy tắc, skill mới, bài học), lập phiếu đề xuất đưa vào kho `D:\CLAUDE` theo skill quản lý kho; chính thức khi đóng dự án.

---

## 8. Bài học

- Trí nhớ về dự án có thể sai (add-in không dùng Cloudflare; API-CLI không phải lõi nhiều nhà cung cấp): đọc mã và tài liệu trước khi kết luận.
- Số liệu gói miễn phí đổi rất nhanh (Gemini CLI ngừng 18/6, Cerebras hết free 16/7, Mistral đổi khoảng 5/11): kiểm lại ngày trước khi dựa vào.
- Gemini free không dùng làm thước đo chất lượng; hết hạn mức thì ghi "Chưa kiểm được".
- Anh ưu tiên phương án tự động; các phương án thủ công nên đưa ra sau cùng và kèm rõ mức thủ công.
- Khi anh trả lời ô tích bằng một câu hỏi mới, coi đó là xin làm rõ, không phải đã chọn.

---

## 9. Nguồn chính (đã dùng trong phiên)

- Điều khoản Thương mại Anthropic: https://www.anthropic.com/legal/commercial-terms
- Điều khoản cá nhân (bản lưu trữ, cấm chia sẻ khóa API): https://www.anthropic.com/legal/archive/1c3ea439-f9b9-435d-9180-3148ab923fbe
- Điều khoản cá nhân (truy cập tự động): https://www.anthropic.com/legal/consumer-terms
- Workspace: https://platform.claude.com/docs/en/manage-claude/workspaces.md
- Admin API: https://platform.claude.com/docs/en/manage-claude/admin-api.md
- Quản lý workspace (hỗ trợ): https://support.anthropic.com/en/articles/9796807-creating-and-managing-workspaces
- Xử lý sự cố Gemini API: https://ai.google.dev/gemini-api/docs/troubleshooting
- So sánh API miễn phí: https://openrouter.ai/blog/tutorials/free-llm-apis-compared/ · https://apidog.com/blog/free-ai-api/ · https://gravity.fast/data/free-llm-api-tiers/
- Gemini CLI ngừng gói free: https://www.tembo.io/blog/gemini-cli-pricing
- OpenClaw zero-token: https://raw.githubusercontent.com/linuxhsj/openclaw-zero-token/main/README.md
- Chat không đăng nhập: https://www.gradually.ai/en/ai-chatbots/ · https://support.google.com/gemini/answer/13275745
- Hộp thoại Office: https://learn.microsoft.com/office/dev/add-ins/develop/dialog-best-practices

---

## 10. Cập nhật 18/11/2026 (tối): "Chuyển tay nhưng tự động hoá"

**Kết luận:** làm được ở mức bán tự động, không cần máy chủ; nhưng mất vòng lặp tác tử (lõi add-in). Chỉ hợp làm chế độ phụ "Hỏi qua chat web" cho việc một lượt.

**Ba mức** [Nguồn ngoài + suy luận, chưa thử trên máy anh]
| Mức | Cách | Thao tác/lượt | Đánh đổi |
|---|---|---|---|
| 1 Bán tự động | Add-in chép lời nhắc vào bộ nhớ tạm + mở trang chat; anh Ctrl+V, Enter, bấm Copy; add-in tự đọc bộ nhớ tạm (hoặc 1 nút Dán), xem trước, áp dụng | 1 dán + 3 bấm | Không thành phần ngoài; rủi ro điều khoản thấp nhất |
| 2 Tự động trên Word web | Tiện ích Edge/Chrome chạy trong khung add-in và tab chat, tự dán/gửi/lấy về | Chỉ duyệt | Chỉ Word trên web (Word cài máy dùng WebView2, tiện ích không vào được — chưa kiểm); trái điều khoản tài khoản cá nhân |
| 3 Tự động trên Word cài máy | Mức 2 + chương trình cầu nối cài máy | Chỉ duyệt | Trái quy tắc "không máy chủ riêng"; rủi ro như mức 2 |

Nạp lời nhắc qua địa chỉ `?q=` chỉ dùng cho lời nhắc ngắn: Claude cắt ~14.000 ký tự (support.claude.com/en/articles/14729294); ChatGPT giới hạn ngắn và tự gửi (simonwillison.net/b/8625); Gemini không hỗ trợ chính thức. Lượt Word 11–35 nghìn token → dùng bộ nhớ tạm.

**Hạn chế tính năng** [Đã đọc mã: `loop.ts` tối đa 20 bước/lượt; 42 công cụ: 23 Word, 10 Excel, 9 PPT]
- AI tự đọc/tìm tài liệu (`word_find`, `read_paragraphs`, `get_outline`, `read_clause`, `read_tables`, `excel_read_range`…): **mất**, add-in phải gói sẵn ngữ cảnh.
- Sửa Word (thay, sửa đoạn, chèn/xoá, định dạng, điều khoản, điền bảng): **giữ** qua khối lệnh cố định + xem trước; kém chắc hơn.
- Bình luận, sửa đổi theo dõi: **giữ** nếu gửi kèm mã.
- Việc nhiều bước có tự kiểm (điền mẫu, rà nhất quán, dựng slide rồi kiểm): **giảm mạnh/mất**.
- Skill: **giảm** (không nạp `references/` theo nhu cầu).
- Chọn model, Suy nghĩ kỹ, tự thử lại, nén lịch sử: **mất**.
- Đính kèm PDF/ảnh: **mất** khi không đăng nhập.
- Cấp cho công ty: **không hợp**.
- Giữ tốt: việc một phát (viết lại đoạn chọn, góp ý điều khoản, dịch, soát lỗi).

Lối tự động thật đã có: đăng nhập OpenRouter (`oauth.ts`) lấy key một bấm, giữ đủ tính năng; hạn mức model free thấp.

**Quyết định:** anh chọn **làm bài đo kỹ thuật Mức 1** trước khi viết mã; không làm Mức 2–3.

**Bài đo Mức 1 — tiêu chí nghiệm thu**
| # | Điểm đo | Đạt khi |
|---|---|---|
| 1 | Đọc bộ nhớ tạm trong khung add-in (Word cài máy và Word web) | `navigator.clipboard.readText()` trả được chữ, hoặc ghi rõ lỗi quyền → dùng nút Dán dự phòng |
| 2 | Ghi bộ nhớ tạm + mở trang chat từ khung (`writeText`, `Office.context.ui.openBrowserWindow` / `window.open`) | Trang chat mở ở trình duyệt, lời nhắc dán được |
| 3 | Khối lệnh còn nguyên sau khi chép bằng nút Copy của Claude / ChatGPT / Gemini | Bộ đọc của add-in phân tích đúng 3/3 nguồn trên một bài mẫu |
Model · effort: Opus 5.5 · Low, Claude Code mở ở `D:\ADD IN`, chỉ ở chế độ thử, không sửa đường chính.

Câu mở đầu chat mới: *"Đọc `D:\ADD IN\tasks\ban-giao-phan-phoi-cong-ty-va-che-do-web.md` mục 10, làm bài đo Mức 1."*

---

## 11. Quyết định 18/11/2026 23:50: bỏ hẳn chuyển tay và chế độ web

- **Bỏ** mọi mức chuyển tay / chat web (Mức 1–3 ở mục 10, cả bài đo Mức 1). Lý do: mất vòng lặp tác tử và nhóm công cụ đọc tài liệu; phần giữ được không đáng công làm.
- **Tập trung hướng API nhiều nguồn.**
- Hiện trạng [Đã đọc mã]: `providers.ts` có 17 nguồn (openrouter, gemini, groq, anthropic, openrouter-claude, deepseek-anthropic, deepseek-flash, deepseek-v4-pro, minimax, kimi, openai, mistral, cerebras, deepseek, sambanova, nvidia, cohere). **Chưa có tự chuyển nguồn khi lỗi**: `errors.ts` dòng 166 chỉ nhắc người dùng tự đổi nguồn rồi gửi lại.
- Ứng viên việc tiếp theo (chờ anh chọn): (a) tự chuyển nguồn dự phòng khi 429/503/im lặng; (b) đo chất lượng tiếng Việt pháp lý và giá của DeepSeek/MiniMax/Kimi/Groq/OpenRouter trên bộ bài cố định; (c) cấp key công ty P1 (workspace Console) hoặc P2 (cổng key); (d) bài đo chẩn đoán Gemini (mục 7 #2).

---

## 12. Kế hoạch (CHỜ DUYỆT) — hai chế độ "Miễn phí" / "Trả phí" và xoay vòng nguồn miễn phí

Yêu cầu của anh (18/11/2026 23:55): đặt 2 chế độ chung cho các nguồn. Chế độ Miễn phí chỉ hiện nhà cung cấp có API miễn phí; chế độ Trả phí chỉ hiện nguồn thu phí. Xoay vòng chỉ áp dụng cho nguồn miễn phí. Thử chế độ Trả phí chỉ dùng key Claude anh đã nhập sẵn.

**Căn cứ mã [Đã đọc mã]:** `Provider` trong `providers.ts` chưa có trường miễn phí/trả phí. `buildStepRequest` (`llm.ts`) dựng yêu cầu riêng cho 3 kiểu gọi (openai/gemini/anthropic) từ cùng một lịch sử, có tùy chọn `plain` bỏ khối riêng của từng nguồn → đổi nguồn giữa chừng khả thi. `errors.ts` đã phân biệt 429 theo ngày (`DAILY_QUOTA`), đọc `retry-after`.

**Phân loại nguồn (ban đầu) [Chưa kiểm trang chính thức; kiểm lại ngày làm]**
| Chế độ | Nguồn |
|---|---|
| Miễn phí | Google Gemini (AI Studio), Groq, OpenRouter (chỉ model đuôi `:free`) |
| Trả phí | Anthropic, OpenRouter (model trả phí) + OpenRouter·Claude, DeepSeek (3 kênh), MiniMax, Kimi, OpenAI |
| Chưa xếp | Mistral, Cohere (nguồn mâu thuẫn), Cerebras (hết free thường trực 16/7/2026); SambaNova, NVIDIA bị chặn CORS |
| Nguồn tự khai | Người dùng tự chọn chế độ khi khai |

**Các bước**
1. Thêm trường `tier: "free" | "paid" | "both"` (OpenRouter = both, lọc model theo `:free`) và công tắc chế độ trong Cài đặt + nhãn model; mặc định Miễn phí.
2. Bộ chọn nguồn/model lọc theo chế độ đang bật.
3. Xoay vòng (chỉ chế độ Miễn phí): thứ tự mặc định, sửa được; gặp 429 theo phút / 503 / im lặng 150 giây → chuyển nguồn kế tiếp, gọi lại bằng `plain`; 429 theo ngày → đánh dấu nguồn "hết hôm nay" đến mốc đặt lại; báo trên khung đang dùng nguồn nào. Hết mọi nguồn → dừng và báo, **không bao giờ tự chuyển sang trả phí**.
4. Kiểm thử máy (`npm test`) cho lọc chế độ và xoay vòng (giả lập lỗi).
5. Thử thật: chế độ Trả phí chỉ với key Claude có sẵn; chế độ Miễn phí với các key free đang có.

**Tiêu chí nghiệm thu**
| # | Tiêu chí | Bằng chứng cần có |
|---|---|---|
| T1 | Chế độ Miễn phí chỉ hiện nguồn/model miễn phí; Trả phí chỉ hiện nguồn thu phí | Kiểm thử máy + ảnh chụp 2 chế độ |
| T2 | 429/503/im lặng ở nguồn free → tự chuyển nguồn kế tiếp, lượt làm xong | Kiểm thử giả lập lỗi xanh; log lượt thật |
| T3 | 429 theo ngày → nguồn bị bỏ qua đến mốc đặt lại | Kiểm thử máy |
| T4 | Không bao giờ tự gọi nguồn trả phí từ chế độ Miễn phí | Kiểm thử máy |
| T5 | Chế độ Trả phí chạy đúng 1 bài Word bằng key Claude có sẵn | JSON kết quả trong `tasks/dot-9-bang-chung/` |
| T6 | Toàn bộ kiểm thử cũ vẫn xanh; không lộ key | `npm test`; rà log |
Model · effort: Opus 5.5 · Medium, Claude Code mở ở `D:\ADD IN`.

**Đã duyệt (19/11/2026 ~00:00):** anh duyệt kế hoạch mục 12, làm trong Claude Code. Thứ tự xoay vòng mặc định: **Gemini → Groq → OpenRouter `:free`** (sửa được trong Cài đặt).

**Kết quả kiểm Mistral, Cohere (19/11/2026):**
| Nguồn | Kết quả | Căn cứ | Xếp |
|---|---|---|---|
| Mistral | Gói Free không có API; API cần bật thanh toán. Gói Pro kèm 15 USD/tháng tín dụng API | Trang chính thức mistral.ai/pricing (đọc 19/11/2026); agentdeals.dev: từ khoảng 5/11/2026 Free không còn API. Một số trang cũ vẫn ghi "Experiment" → lỗi thời | **Trả phí** |
| Cohere | Key dùng thử miễn phí, giới hạn tốc độ; **không được dùng cho mục đích sản xuất hoặc thương mại** | Trang chính thức cohere.com/pricing; docs.cohere.com/docs/going-live: Chat tối đa 1.000 lần gọi/tháng với key dùng thử. Trang tổng hợp: 20 yêu cầu/phút | **Không đưa vào Miễn phí** (dùng cho công việc công ty = thương mại; 1.000 lần/tháng quá ít cho vòng lặp tới 20 bước). Có key sản xuất thì xếp Trả phí |
Cerebras: giữ ẩn (hết free thường trực 16/7/2026) cho đến khi kiểm lại.

**Bài học:** `device_commit_files` dùng lại cùng `stagedPath` có thể ghi nội dung cũ dù báo "written" → mỗi lần ghi dùng tên tệp tạm mới và kiểm lại kích thước sau khi ghi.

Câu mở đầu cho Claude Code (mở ở `D:\ADD IN`): *"Đọc `tasks/ban-giao-phan-phoi-cong-ty-va-che-do-web.md` mục 12 (đã duyệt) và `CLAUDE.md`. Lập `tasks/dot-9.md` theo mục 12, rồi làm đợt 9: hai chế độ Miễn phí/Trả phí, xoay vòng nguồn miễn phí. Báo bảng T1–T6 có bằng chứng."*

**Sửa quyết định (19/11/2026 00:03, anh Dũng):** đưa **Cohere vào chế độ Miễn phí**. Lý do: add-in phân phối cho cá nhân, ai dùng thì tự nhập key của người đó. Thứ tự xoay vòng mặc định: **Gemini → Groq → OpenRouter `:free` → Cohere** (Cohere cuối vì chỉ 1.000 lần gọi/tháng). Ghi chú hiện ở Cài đặt khi chọn Cohere: "Key dùng thử: miễn phí, giới hạn tốc độ, Cohere không cho dùng vào mục đích thương mại." Cohere kiểu gọi `openai` qua `api.cohere.ai/compatibility/v1` — cần kiểm gọi công cụ chạy được khi làm đợt 9.
