# Bàn giao — AI Đa Nguồn

Cập nhật 24/11/2026. Giữ tệp này **dưới 15 KB**: mỗi phiên sửa tại chỗ, gạch hoặc xóa mục đã xong. Chi tiết cũ vào `tasks/luu/`, chi tiết đợt vào tệp đợt, bài học vào `docs/bai-hoc.md`.

- Bản đầy đủ tới 24/11/2026 (mọi quyết định nguyên văn, diễn biến từng đợt): `tasks/luu/HANDOFF-den-2026-11-24.md`; trước đó `tasks/luu/HANDOFF-den-2026-11-17.md`.
- Bài học (Office.js, WebView2, bài thử, quy trình): `docs/bai-hoc.md`.

## 1. Trạng thái

- **Đang chạy: v4.11.11** (phát hành 24/11/2026, `main@ab0cae4` → `gh-pages@b5eb0a5`; Tự kiểm tra Word 41/41 trên bản phát hành; kiểm thử máy 886/886).
- **Đang làm:** không có việc mã dở. Dọn tài liệu và thư mục (24/11/2026) đã gộp vào `main` trên máy, **chưa đẩy lên GitHub**: anh chọn gọn tệp này trước rồi đẩy một lần.

## 2. Phiên bản gần đây

| Bản | Ngày | Nội dung | Chi tiết |
|---|---|---|---|
| 4.11.10 + 4.11.11 | 24/11 | Hai buổi thử model miễn phí soạn văn bản: kiểm ngưỡng luật (`legalcheck.ts`), kiểm thứ / từ lạ / luật hết hiệu lực / sai dấu (`textcheck.ts`), loại `giay_uy_quyen`, nghị quyết, bên cá nhân ký "BÊN B", cảnh báo "kể đã sửa A → B mà không ghi". Điểm TB: Nemotron 5,7 → 8,3; Cohere 5,5 → 7,5; dots 4,7 → 6,5. Soát mã độc lập: sửa hết 3 cao + 8 TB | `tasks/viec-le/thu-mien-phi-vong3/`, `thu-mien-phi-hdcn/`, `thu-mien-phi-vbhc/` |
| 4.11.9 | 23/11 | "Điều N" không khớp nhầm mục "N." của phụ lục đánh số lại | `tasks/viec-le/va-tra-dieu-khoan/` |
| 4.11.8 | 23/11 | Soạn hợp đồng đúng thể thức (`hop_dong`, kỹ năng `soan-hop-dong`), đọc đoạn báo ô bảng, chặn xóa vắt mép bảng | `tasks/viec-le/sua-23-11/` |
| 4.11.3 – 4.11.7 | 22–23/11 | Đợt 12: đo token đầu vào (A), Excel ngày dd/mm (B), sửa theo soát mã (C), .docx đính kèm đọc theo điều (D), bộ nhớ dài hạn (E) | `tasks/dot/dot-12/` |
| 4.11.1, 4.11.2 | 22/11 | Bảng PowerPoint tự co vừa slide, Miễn phí xoay vòng OpenRouter trước | `tasks/viec-le/sua-22-11/` |

Các bản cũ hơn (đợt 0A → 11, bản 3.6 → 4.11.0): bảng "Đã xong" trong `tasks/luu/HANDOFF-den-2026-11-24.md`; mỗi đợt có thư mục `tasks/dot/dot-<N>/`.

## 3. Quyết định còn hiệu lực

Nguyên văn và lý do đầy đủ: mục "Quyết định và lý do" của bản lưu 24/11. Quy tắc kỹ thuật rút từ các quyết định này nằm ở `docs/quy-tac/`.

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
- 24/11: CLAUDE.md chỉ bắt đọc tệp này; quy tắc theo phần ở `docs/quy-tac/`; `tasks/` gom theo `dot/`, `viec-le/`, `luu/`.

## 4. Vấn đề mở

### Chưa kiểm được vì thiếu tiền hoặc key (cần anh)

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

### Giới hạn đã biết, chưa định sửa

- Word: không đọc số trường LISTNUM / AUTONUM; điều cuối kéo tới hết văn bản; chưa bấm được theo "Điều 5.3"; so sánh phiên bản để lại tài liệu ẩn tới khi thoát Word; cửa sổ Word thu nhỏ thì không chèn được nhận xét.
- PowerPoint: lần chèn biểu đồ đầu trên bài mới để lại một bố cục thừa; không có API ghi chú diễn giả, số liệu biểu đồ.
- OpenRouter: không thu hồi được khóa, mỗi lần đăng nhập cấp thêm một khóa; sau Đăng xuất rồi đăng nhập lại, nguồn đang chọn không tự quay về OpenRouter.
- Nguồn có danh sách model cố định (MiniMax, Kimi…): key chỉ được kiểm ở lượt gọi đầu.

### Chờ anh quyết hoặc làm

- Đẩy `main` (dọn tài liệu 24/11) lên GitHub.
- 8 worktree còn commit chưa gộp vào `main` (anh xét giữ hay bỏ): `claude-addon-comparison-test-2f7a5b` (39), `mo-bao-loi-an-word-2c4edb` (48), `nice-grothendieck-6cd46e` = `tham-do-antigravity` (3), `openrouter-excel-powerpoint-test-cbd6be` (39), `optimize-add-in-testing-ea9a16` = `thu-mien-phi-vong-3` (3), `project-architecture-audit-38106c` (1), `test-ai-land-transfer-contract-6c607e` = `thu-mien-phi-van-ban` (1), `two-questions-old-session-3c4e5a` (42). Hai nhánh vòng thử miễn phí đã phát hành bằng commit khác nên có thể chỉ còn commit trùng nội dung — kiểm trước khi bỏ.
- Hai thư mục mồ côi `dot-0a-office-addon-compare-47836b`, `dot-4-a4b5e7` trong `.claude\worktrees\`: đưa vào Thùng rác.
- Xóa các khóa "AI Da Nguon (Office)" thừa ở `openrouter.ai/settings/keys`.
- Nguồn Gemini trên máy còn nhớ "Gemini 2.5 Flash" (Google ngừng cấp cho người dùng mới): chọn lại Gemini 3.8 Flash.
- Thăm dò gọi model Google qua Antigravity (ý của anh 23/11): kết quả ở nhánh `claude/tham-do-antigravity`, chưa xét.

## 5. Bước tiếp theo

1. Đẩy `main` lên GitHub (chỉ tài liệu và lời chú thích; Actions phát hành lại cùng mã).
2. Xét 8 worktree chưa gộp và nhánh Antigravity (mục 4).
3. **Buổi kiểm bù** khi anh đã nạp tiền hoặc có key: chạy hết nhóm "chưa kiểm được" ở mục 4.
4. Việc tùy chọn (mục D của `tasks/dot/dot-5/dot-5.md`), anh chọn: tải bản ghi hội thoại, gợi ý gạch chân bấm để chấp nhận trong Word, connectors (MCP), công cụ chạy mã Office.js do AI viết, phím tắt mở khung, đọc một trang web cụ thể, Outlook, Sign in with Poe, hàm `=AI(…)` trong ô Excel. Đã làm: điền mẫu (Đợt 8), bộ nhớ dài hạn (Đợt 12 E).

Câu mở đầu phiên sau: `Đọc tasks/HANDOFF.md mục Bước tiếp theo rồi làm tiếp.`

## 6. Bài học mới của phiên

- (24/11/2026) Tệp bàn giao lại phình tới 87 KB sau lần viết gọn 18/11, vì "giữ nguyên văn mục Quyết định và Bài học" cộng với mỗi phiên chỉ thêm. Nay Bài học có tệp riêng (`docs/bai-hoc.md`), tệp này có trần 15 KB ghi ở đầu và ở CLAUDE.md.
