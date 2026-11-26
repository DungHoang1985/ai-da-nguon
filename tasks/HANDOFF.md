# Bàn giao — AI Đa Nguồn

Cập nhật 26/11/2026. Giữ tệp này **dưới 15 KB**: mỗi phiên sửa tại chỗ, gạch hoặc xóa mục đã xong. Chi tiết cũ vào `tasks/luu/`, chi tiết đợt vào tệp đợt, bài học vào `docs/bai-hoc.md`.

- Bản đầy đủ tới 24/11/2026 (mọi quyết định nguyên văn, diễn biến từng đợt): `tasks/luu/HANDOFF-den-2026-11-24.md`; trước đó `tasks/luu/HANDOFF-den-2026-11-17.md`.
- Bài học (Office.js, WebView2, bài thử, quy trình): `docs/bai-hoc.md`.

## 1. Trạng thái

- **Đang chạy: v4.13.1** (phát hành 25/11/2026, `main@bd6dea7` → `gh-pages@9f4a154`; Tự kiểm tra trên bản phát hành 26/11: Word 43/43, Excel 43/43, PowerPoint 27/27; kiểm thử máy 961/961). Trước đó: **v4.13.0** (`main@f4f57d2` → `gh-pages@000bfd6`).
- 4.13.1 = cổng soát lập luận bằng mã: lệnh ghi câu chữ Word (`needsExplain`) lần đầu trong lượt mà chưa gọi `word_explain_edits` thì không ghi, nhắc model soát trước (một lần mỗi lượt). Chạy thật Miễn phí (Groq / Cohere): AI soát trước khi ghi 2/2 lần.
- 26/11, sửa bài thử (không thuộc bản add-in): `mo-office.ps1` kết thúc PowerPoint treo sau Quit (lỗi lần 3); `cdp.mjs eval` chờ khung vẽ xong, trắng thì tự tải lại một lần.

## 2. Phiên bản gần đây

| Bản | Ngày | Nội dung | Chi tiết |
|---|---|---|---|
| 4.13.1 | 25/11 | Cổng bằng mã: lệnh ghi câu chữ Word lần đầu trong lượt phải soát lập luận trước (nhắc một lần; sửa máy móc thì gọi lại). Chạy thật Miễn phí: AI soát trước khi ghi 2/2 lần. Kiểm thử máy 961/961 | `tasks/viec-le/rea-claude-chinh-thuc/` |
| 4.13.0 | 25/11 | Công cụ soát lập luận Word `word_explain_edits`: AI nêu căn cứ có loại nguồn cho từng sửa đổi câu chữ pháp lý; người soát (lượt model riêng, không thấy hội thoại) đối chiếu văn bản, sửa đổi chờ duyệt, chỉ dẫn riêng, kỹ năng đã nạp; căn cứ chưa xác nhận → AI phải sửa đề xuất hoặc hỏi. Mốc chặn phình: Trả phí 33.950, Miễn phí 20.800 ký tự. Kiểm thử máy 954/954, Tự kiểm tra trên bản phát hành Word 43/43, Excel 43/43, PowerPoint 27/27 | `tasks/viec-le/rea-claude-chinh-thuc/` |

Các bản cũ hơn (đợt 0A → 12, bản 3.6 → 4.12.0): bảng "Đã xong" trong `tasks/luu/HANDOFF-den-2026-11-24.md`; mỗi đợt có thư mục `tasks/dot/dot-<N>/`.

## 3. Quyết định còn hiệu lực

Nguyên văn và lý do đầy đủ: mục "Quyết định và lý do" của bản lưu 24/11. Quy tắc kỹ thuật rút từ các quyết định này nằm ở `docs/quy-tac/`.

- 25/11: lời nhắc dặn **đọc lại sau khi ghi** (một lần cho mọi chỗ, cuối việc) thay quy tắc cũ "KHÔNG ĐỌC LẠI" (anh quyết, chấp nhận thêm lần gọi model để báo đúng).
- AI sửa thẳng vào tài liệu + Hoàn tác; Word luôn ghi Track Changes (đã bỏ công tắc tắt); thao tác rủi ro luôn hỏi, "Cho phép hết lần này" không áp dụng cho thao tác rủi ro.
- Kỹ năng chính thức chuyển thể, không chép nguyên văn; không ghi số hiệu văn bản pháp luật. Tự chọn kỹ năng bằng từ khóa không dấu + danh mục để model tự gọi `use_skill`.
- Đăng nhập: OpenRouter một lần bấm + tín dụng API gói Max. **Không làm** gói Claude qua Agent SDK (vùng xám).
- Bố cục khung chat một dòng đầu, giống bản chính thức.
- Word: đọc số điều bằng `listItemOrNullObject` + mẫu số gõ tay; AI dùng chỉ số đoạn [¶n] + `expect`; Hoàn tác = từ chối đúng các sửa đổi mới xuất hiện (không dùng `Document.undo`); so sánh phiên bản tự so bằng mã theo điều khoản, redline chỉ mở ở cửa sổ mới.
- Khối suy nghĩ lưu trong `raw`, gửi lại nguyên văn; Suy nghĩ kỹ chỉ đổi tham số ngoài lời nhắc; soát lập luận trước khi ghi mặc định tắt.
- PDF quét qua OpenRouter chỉ dùng engine `native` (không OCR tính phí riêng).
- Miễn phí ở OpenRouter ưu tiên Nemotron 3 Ultra free; ERNIE 5.5 Flash free mang dấu ⚠.
- Excel: tô sáng ô và sheet "Nhật ký AI" mặc định tắt.
- Kiểm thử phải bật phần mềm thì chạy trên desktop ẩn.
- Nút Hoàn tác chỉ hiện số thao tác (không hiện lý do) — để nguyên.
- Không thêm Z.ai, DashScope quốc tế (máy chủ chặn gọi từ khung add-in).
- Định dạng kết quả đọc Word: **giữ JSON** (thử A/B 23/11 với dòng gọn: chất lượng ngang nhau, không đo được tiết kiệm token; báo cáo ở nhánh cục bộ `claude/ab-dinh-dang-gon`, không phát hành).
- Tên tài khoản Office của anh còn ở commit cũ `a42347a` của kho công khai: không viết lại lịch sử. Đã đóng.
- Gọi model Google qua Antigravity: **không làm** (điều khoản Google cấm phần mềm bên thứ ba; dùng nguồn Gemini bằng key AI Studio) — `tasks/viec-le/tham-do-antigravity/`.
- 24/11: kết quả so sánh lớn hoặc có lời AI (21–22/11) cất ở `_luu-cuc-bo/` trên máy, ngoài Git (`.gitignore`), không lên kho công khai.
- 24/11: CLAUDE.md chỉ bắt đọc tệp này; quy tắc theo phần ở `docs/quy-tac/`; `tasks/` gom theo `dot/`, `viec-le/`, `luu/`.

## 4. Vấn đề mở

### Chưa kiểm được vì thiếu tiền hoặc key (cần anh)

- 25/11: key Groq trên máy bị Groq từ chối ("Invalid API Key") — cần anh dán key mới.
- OpenRouter trả 400 "Request path cannot be relayed" cho MỌI lệnh gọi model (cả `:free` lẫn Claude) từ 25/11; `/api/v1/key`, `/credits` vẫn 200 (key hợp lệ, không giới hạn, tài khoản còn khoảng 4 USD); key giả ra 401. Câu trả lời có `cf-ray` (đúng là OpenRouter, không phải proxy trên máy), đi qua điểm Cloudflare Hồng Kông (HKG); máy không dùng proxy. Cùng ngày nhiều người ở Trung Quốc, Việt Nam gặp y hệt; một người nói dùng VPN thì chạy. Kết luận (chưa xác nhận chính thức): OpenRouter đang chặn yêu cầu suy luận đi từ vùng này / qua HKG. Add-in không sửa được (không thêm máy chủ riêng). Cách thử: anh bật VPN ra nước khác rồi bấm "Kiểm tra từng nguồn"; hoặc hỏi hỗ trợ OpenRouter.

- Gemini: 5 bài trên hợp đồng thử, PDF quét 3 và 7 trang, Suy nghĩ kỹ (gói miễn phí hết hạn mức rất nhanh; bảng B1, B2 ở `tasks/dot/dot-5/dot-5.md`). Công cụ của Đợt 7 chưa Chạy thật với Gemini.
- DeepSeek (cả hai kênh): hết tiền. MiniMax, Kimi: chưa có key (tiêu chí 4 Đợt 6 cần một nguồn nhóm này có tiền).
- OpenRouter: PDF qua kênh Claude cần số dư từ 0,50 USD; model trả tiền chưa thử.
- OpenAI: lỗi không đọc được chi tiết, nghi số dư hoặc quyền dùng model.
- Model Claude đời cũ dùng `budget_tokens`; ca soát lập luận trả "CHƯA ĐẠT" (Đợt 1B).

### Việc kỹ thuật chưa làm

- Hoàn tác sau Làm tiếp vẫn tính theo từng lượt, không gộp. Kết quả `word_find` đánh dấu chỗ khớp bằng «…», có lần AI tưởng là định dạng văn bản. Dòng "lịch sử đầy" chưa nhìn trong Office.
- Cảnh báo "nói đã sửa mà không ghi" ở lượt Làm tiếp chỉ đọc mới có Kiểm thử máy; hệ số ước token cho Gemini chưa đo riêng; `ppt_edit_table` thêm hàng chưa Chạy thật với model.
- Mục soát mã để lại: Đợt 4 còn 10 mục (`tasks/dot/dot-4/dot-4.md`); Đợt 3 phần A còn 1, phần B còn 6; Đợt 7 chủ yếu mục thấp (`tasks/dot/dot-7/dot-7.md`); mục thấp của soát mã 4.11.10–11 ở báo cáo vòng 3.
- Ghi chéo với hai tài liệu cùng ứng dụng chưa thử trong Office. Desktop ẩn chưa thăm dò ba ứng dụng cùng lúc.
- Bản quét: đọc hai lần độc lập rồi máy so số liệu; máy kiểm dãy Điều – khoản – điểm trên bản chép.
- Groq: hạn token mỗi phút theo model chưa có bảng; chưa đo khi model gọi công cụ không có trong danh sách; đổi Miễn phí → Trả phí giữa hội thoại khi lịch sử có `load_tools`.
- Đợt 6: SiliconFlow, Qwen, ModelScope, StepFun chưa thêm sẵn (dùng qua nguồn tự khai); Bedrock, Foundry, Vercel chưa làm.
- Hạn chờ im lặng 150 giây mới thấy chạy đúng một lần (Gemini, bản cục bộ).

- 4.13.1 để lại: thẻ bước bị cổng chặn hiện như lỗi (⚠️); cổng chỉ bảo đảm AI đã GỌI soát, không bảo đảm AI làm theo kết luận (model yếu vẫn ghi căn cứ chưa kiểm được). 4.13.0 phản biện để lại: người soát chỉ thấy kỹ năng tự chọn đầu lượt (kỹ năng gọi `use_skill` giữa lượt → "chưa nạp nguồn", an toàn nhưng thừa câu hỏi); căn cứ `user_instruction` chỉ đối chiếu chỉ dẫn riêng đã lưu, không thấy lời yêu cầu trong chat (cố ý: người soát không thấy hội thoại); sửa đổi chờ duyệt chưa gắn số đoạn; từ khóa "dieu khoan" làm nhóm `soat` (Miễn phí) bật thường xuyên; người soát dùng trần đầu ra như lượt chính.

### Giới hạn đã biết, chưa định sửa

- Word: không đọc số trường LISTNUM / AUTONUM; điều cuối kéo tới hết văn bản; chưa bấm được theo "Điều 5.3"; so sánh phiên bản để lại tài liệu ẩn tới khi thoát Word; cửa sổ Word thu nhỏ thì không chèn được nhận xét.
- PowerPoint: lần chèn biểu đồ đầu trên bài mới để lại một bố cục thừa; không có API ghi chú diễn giả, số liệu biểu đồ.
- OpenRouter: không thu hồi được khóa, mỗi lần đăng nhập cấp thêm một khóa; sau Đăng xuất rồi đăng nhập lại, nguồn đang chọn không tự quay về OpenRouter.
- Nguồn có danh sách model cố định (MiniMax, Kimi…): key chỉ được kiểm ở lượt gọi đầu.

### Chờ anh quyết hoặc làm

- Tối ưu theo mã add-in chính thức (25/11, báo cáo mục 7): mức 1 xong; mức 2 phần 1 ở 4.12.0; còn việc 6 (công cụ soát lập luận mới, bật mặc định cho Word — anh đã chọn) → 4.13.0; mức 2 (5 việc; việc 5 nay chỉ còn phần cổng tự động "ghi mà chưa đọc lại" vì quy tắc lời nhắc đã đổi), mức 3 (PowerPoint storyboard, biểu đồ theo đặc tả). Bản trích gốc lời nhắc chỉ ở `_luu-cuc-bo/rea-claude-office-2026-11-25/` (ngoài Git, không đưa lên kho công khai).

- Rà soát kiến trúc 20/11, bảng "Phát hiện để lại" (`tasks/viec-le/ra-soat-kien-truc/ra-soat-kien-truc-2026-11-20.md`): nhóm 1 (B1, B2, B3, B4, B6, B23) xong ở 4.11.13. Còn: nhóm 2 kiến trúc (B7 tách `Chat.tsx`, B8 sổ công cụ một nguồn, B9 `quirks`, B13 tách `wordTools.ts`) — việc lớn, cần kế hoạch riêng; nhóm 3 hiệu năng (B10 gói 1,26 MB, B15 lint, B16, B18); nhóm 4 bảo mật nhỏ (B5, B11, B12, B17, B19, B20, B21, B22). B14 giữ nguyên.
- Hai thư mục mồ côi `dot-0a-office-addon-compare-47836b`, `dot-4-a4b5e7` trong `.claude\worktrees\`: đưa vào Thùng rác.
- Xóa các khóa "AI Da Nguon (Office)" thừa ở `openrouter.ai/settings/keys`.
- Nguồn Gemini trên máy còn nhớ "Gemini 2.5 Flash" (Google ngừng cấp cho người dùng mới): chọn lại Gemini 3.8 Flash.

## 5. Bước tiếp theo

1. 4.13.1 xong (Tự kiểm tra đạt; Chạy thật Trả phí 26/11, Anthropic · Claude Sonnet 5.5: AI tự soát trước khi ghi, căn cứ chưa kiểm được nói rõ; riêng không đọc lại sau khi ghi — `tasks/viec-le/rea-claude-chinh-thuc/` mục 9). Chọn tiếp: mức 2 của báo cáo, hoặc nhóm 3 rà soát kiến trúc (khung mở nhanh hơn), hoặc lập kế hoạch nhóm 2.
2. **Buổi kiểm bù** khi anh đã nạp tiền hoặc có key: chạy hết nhóm "chưa kiểm được" ở mục 4.
3. Việc tùy chọn (mục D của `tasks/dot/dot-5/dot-5.md`), anh chọn: tải bản ghi hội thoại, gợi ý gạch chân bấm để chấp nhận trong Word, connectors (MCP), công cụ chạy mã Office.js do AI viết, phím tắt mở khung, đọc một trang web cụ thể, Outlook, Sign in with Poe, hàm `=AI(…)` trong ô Excel. Đã làm: điền mẫu (Đợt 8), bộ nhớ dài hạn (Đợt 12 E).

Câu mở đầu phiên sau: `Đọc tasks/HANDOFF.md mục Bước tiếp theo rồi làm tiếp.`

## 6. Bài học mới của phiên

- (25/11/2026) Gói `pivot.claude.ai/m-addin` (9,8 MB, một tệp) làm REA `analyze_javascript_application` chạm trần trạng thái bí danh; `recover_javascript_sources` cần Linux. Cách chạy được: phân tích cú pháp bằng `acorn` (không chạy mã) và trích chuỗi dài — lời nhắc hệ thống từng ứng dụng nằm nguyên văn trong gói. Script ở `_luu-cuc-bo/rea-claude-office-2026-11-25/strings.cjs`.

- (24/11/2026) Tệp bàn giao lại phình tới 87 KB sau lần viết gọn 18/11, vì "giữ nguyên văn mục Quyết định và Bài học" cộng với mỗi phiên chỉ thêm. Nay Bài học có tệp riêng (`docs/bai-hoc.md`), tệp này có trần 15 KB ghi ở đầu và ở CLAUDE.md.
- (24/11/2026) Một nhánh sửa lỗi thật (rà soát kiến trúc 20/11, 8 chỗ sửa + 125 dòng kiểm thử) nằm quên 4 ngày ở worktree vì HANDOFF ghi "chưa gộp" lẫn trong đoạn dài. Cuối mỗi phiên: `git branch --no-merged main` và liệt kê mọi nhánh chưa gộp ở mục "Chờ anh quyết".
