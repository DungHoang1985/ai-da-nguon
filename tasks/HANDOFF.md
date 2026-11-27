# Bàn giao — AI Đa Nguồn

Cập nhật 26/11/2026. Giữ tệp này **dưới 15 KB**: mỗi phiên sửa tại chỗ, gạch hoặc xóa mục đã xong. Chi tiết cũ vào `tasks/luu/`, chi tiết đợt vào tệp đợt, bài học vào `docs/bai-hoc.md`.

- Bản đầy đủ tới 24/11/2026 (mọi quyết định nguyên văn, diễn biến từng đợt): `tasks/luu/HANDOFF-den-2026-11-24.md`; trước đó `tasks/luu/HANDOFF-den-2026-11-17.md`.
- Bài học (Office.js, WebView2, bài thử, quy trình): `docs/bai-hoc.md`.

## 1. Trạng thái

- **Đang chạy: v4.17.0** (phát hành 27/11/2026, `main@6be2cf9` → `gh-pages@3cee986`; Tự kiểm tra trên bản phát hành: Word 43/43, Excel 44/44, PowerPoint 27/27; kiểm thử máy 1.100/1.100). Trước đó: **v4.16.0** (`main@e1da897` → `gh-pages@8d99a3d`).
- 4.13.2 = mức 2 việc 5 phần Excel: `excel_format_range` trả chữ hiển thị (`display_preview`) và cảnh báo ô số có thể hiện "####" (ước theo cỡ chữ vì Office.js không trả "####" — đã đo) → coi như đã đọc lại. Chạy thật Trả phí (Claude Sonnet 5.5): AI thấy cảnh báo, tự giãn cột, không còn ô "####" (`tasks/viec-le/rea-claude-chinh-thuc/` mục 10).
- Nhóm 2 rà soát kiến trúc (`tasks/viec-le/ra-soat-kien-truc/ke-hoach-nhom-2.md`). **Nhóm 2 xong, phát hành ở 4.15.0** (gộp theo ý anh): bước 1 B9 `quirks`; 2a/2b tách `wordTools.ts` (2.482 → 853 dòng: `wordCore`, `wordRead`, `wordEdit`, `wordComments`); 4a/4b phần thuần của lượt ra `core/turnPrep.ts` và `prepareFreeRotation`, B5 chụp tài liệu chỉ khi có khung khác mở. Golden y nguyên; kiểm thử máy 1.072/1.072; bản cục bộ: Tự kiểm tra Word 43/43, Excel 44/44, PowerPoint 27/27; Chạy thật Trả phí (tệp đính kèm), Miễn phí (OpenRouter), ghi chéo và đọc chéo Excel ↔ Word đạt. Bước 3 (B8): anh quyết không làm.
- 4.13.3 = nhóm 3 rà soát kiến trúc: B10 mã Tự kiểm tra nạp khi bấm nút (gói khung 1.294 → 1.145 KB); B16 bỏ babel-loader + 6 gói thừa; B18 `Rich` / `Linked` ghi nhớ, `onNav` ổn định; B15 `eslint.config.mjs` riêng, `npm run lint` 0 lỗi / 285 cảnh báo Office.js, CI chặn khi có lỗi lint. tác tử phản biện: không lỗi nghiêm trọng, đã làm các điểm nên làm. Để lại: tách công cụ theo ứng dụng (phần còn lại của B10, cần nạp công cụ bất đồng bộ trong `Chat.tsx`).

- **Nhóm 4 bảo mật nhỏ (27/11, phát hành 4.16.0; Tự kiểm tra bản cục bộ 27/11 trên desktop ẩn: Word 43/43, Excel 44/44, PowerPoint 27/27, không có vi phạm CSP, CDP vẫn nối sau B21):** B11 `db.ts` báo xong ở `oncomplete`, `update()` đọc–sửa–ghi một giao dịch (mục lục hội thoại), mở IndexedDB lỗi → `FallbackKV` chuyển bộ nhớ tạm và khung chat báo; B17 CSP chèn lúc build (`csp-plugin.js`, mã băm script nội tuyến tính trên HTML đã rút gọn); B19 `tracked()` đòi WordApi 1.6; B20 bỏ đường dẫn máy trong kho + kiểm thử máy chặn tái phát, kế hoạch nội bộ chuyển ra `_luu-cuc-bo/ke-hoach-noi-bo/`; B21 `--remote-allow-origins` chỉ 127.0.0.1/localhost; B22 nguồn tự khai nạp ở hàm dựng `Storage` và sau khôi phục. B12: chỉ kho `ai-da-nguon` bật Pages (kiểm 27/11) — không bật Pages cho kho khác. Kiểm thử máy 1.086/1.086, lint 0 lỗi. Phản biện (tác tử phụ): đã sửa lỗi `fn` ném trong `update`, bỏ coi `onblocked` là hỏng, build dừng khi trang thiếu `<meta charset>`; để lại: CSP còn cho `*.microsoft.com` và `connect-src *` (cần thu hẹp sau khi đo trong Office), `untracked()` của wordAdmin (định dạng trang, không thuộc Hoàn tác) và `insertComment` không chặn dưới 1.6 (cố ý). Kiểm thêm dưới CSP: trang đăng nhập (mã nội tuyến, Office.js), worker pdf.js, đính kèm PDF đều đạt.

## 2. Phiên bản gần đây

| Bản | Ngày | Nội dung | Chi tiết |
|---|---|---|---|
| 4.17.0 | 27/11 | Mức 3 PowerPoint: storyboard (key_message, đọc lại tiêu đề), soát văn phong / cỡ chữ / tiêu đề bằng mã trong kết quả công cụ (`pptStyle.ts`), biểu đồ theo đặc tả (unit, source, highlight). Chặn lỗi `\b` sát chữ có dấu bằng kiểm thử máy | `tasks/viec-le/rea-claude-chinh-thuc/` mục 6–7 |
| 4.16.0 | 27/11 | Nhóm 4 bảo mật nhỏ: B11 lưu hội thoại an toàn, B17 CSP, B19 Word dưới 1.6 không ghi, B20 bỏ đường dẫn máy, B21, B22 | `tasks/viec-le/ra-soat-kien-truc/ra-soat-kien-truc-2026-11-20.md` |
| 4.15.0 | 26/11 | Nhóm 2: B9 quirks, tách wordTools.ts, phần thuần của lượt ra lõi, B5. Golden y nguyên | `tasks/viec-le/ra-soat-kien-truc/ke-hoach-nhom-2.md` |
| 4.13.4 | 26/11 | Nhóm 2 bước 0: vá lệnh ghi chéo chen ngang; golden 75 ca giữ nguyên thứ gửi model | `tasks/viec-le/ra-soat-kien-truc/ke-hoach-nhom-2.md` |
| 4.13.3 | 26/11 | Nhóm 3 rà soát kiến trúc: gói khung −148 KB (Tự kiểm tra tách riêng), bỏ babel, ghi nhớ khung chat, lint chạy được và vào CI | `tasks/viec-le/ra-soat-kien-truc/` |
| 4.13.2 | 26/11 | Excel: lệnh định dạng trả chữ hiển thị và cảnh báo ô số bị "####" (ước theo số đo tự giãn cột). Kiểm thử máy 967/967, Tự kiểm tra Excel 44/44 | `docs/quy-tac/excel.md` |

Các bản cũ hơn (đợt 0A → 12, bản 3.6 → 4.12.0): bảng "Đã xong" trong `tasks/luu/HANDOFF-den-2026-11-24.md`; mỗi đợt có thư mục `tasks/dot/dot-<N>/`.

## 3. Quyết định còn hiệu lực

Nguyên văn và lý do đầy đủ: mục "Quyết định và lý do" của bản lưu 24/11. Quy tắc kỹ thuật rút từ các quyết định này nằm ở `docs/quy-tac/`.

- 26/11: **làm liền nhiều bước rồi mới phát hành một lần** (mỗi bước vẫn kiểm đủ trên bản cục bộ và commit trên máy; hỏi anh một lần khi đẩy `main`). Nguồn AI chỉ OpenRouter + Anthropic (dự phòng Cerebras, Cohere).
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

### Nguồn AI (anh chốt 26/11/2026)

- Chỉ dùng OpenRouter (kèm kênh Claude qua OpenRouter) và Anthropic; giữ Cerebras, Cohere làm dự phòng Miễn phí. Đã xóa khỏi add-in key Gemini, Groq, DeepSeek, OpenAI (đều lỗi). "Kiểm tra từng nguồn" 26/11: 5/5 nguồn còn key đều đạt. Không ghi các nguồn đã xóa vào "chưa kiểm được" nữa.
- Còn để biết: OpenRouter — PDF qua kênh Claude cần số dư từ 0,50 USD, model trả tiền chưa thử; model Claude đời cũ dùng `budget_tokens`, ca soát lập luận trả "CHƯA ĐẠT" (Đợt 1B).

### Việc kỹ thuật chưa làm

- Hoàn tác sau Làm tiếp vẫn tính theo từng lượt, không gộp. Kết quả `word_find` đánh dấu chỗ khớp bằng «…», có lần AI tưởng là định dạng văn bản. Dòng "lịch sử đầy" chưa nhìn trong Office.
- Cảnh báo "nói đã sửa mà không ghi" ở lượt Làm tiếp chỉ đọc mới có Kiểm thử máy; hệ số ước token cho Gemini chưa đo riêng; `ppt_edit_table` thêm hàng chưa Chạy thật với model.
- Mục soát mã để lại: Đợt 4 còn 10 mục (`tasks/dot/dot-4/dot-4.md`); Đợt 3 phần A còn 1, phần B còn 6; Đợt 7 chủ yếu mục thấp (`tasks/dot/dot-7/dot-7.md`); mục thấp của soát mã 4.11.10–11 ở báo cáo vòng 3.
- Ghi chéo với hai tài liệu cùng ứng dụng chưa thử trong Office. Desktop ẩn chưa thăm dò ba ứng dụng cùng lúc.
- Bản quét: đọc hai lần độc lập rồi máy so số liệu; máy kiểm dãy Điều – khoản – điểm trên bản chép.
- Đổi Miễn phí → Trả phí giữa hội thoại khi lịch sử có `load_tools`: chưa đo.
- Đợt 6: SiliconFlow, Qwen, ModelScope, StepFun chưa thêm sẵn (dùng qua nguồn tự khai); Bedrock, Foundry, Vercel chưa làm.

- 4.13.1 để lại: thẻ bước bị cổng chặn hiện như lỗi (⚠️); cổng chỉ bảo đảm AI đã GỌI soát, không bảo đảm AI làm theo kết luận (model yếu vẫn ghi căn cứ chưa kiểm được). 4.13.0 phản biện để lại: người soát chỉ thấy kỹ năng tự chọn đầu lượt (kỹ năng gọi `use_skill` giữa lượt → "chưa nạp nguồn", an toàn nhưng thừa câu hỏi); căn cứ `user_instruction` chỉ đối chiếu chỉ dẫn riêng đã lưu, không thấy lời yêu cầu trong chat (cố ý: người soát không thấy hội thoại); sửa đổi chờ duyệt chưa gắn số đoạn; từ khóa "dieu khoan" làm nhóm `soat` (Miễn phí) bật thường xuyên; người soát dùng trần đầu ra như lượt chính.

### Giới hạn đã biết, chưa định sửa

- Word: không đọc số trường LISTNUM / AUTONUM; điều cuối kéo tới hết văn bản; chưa bấm được theo "Điều 5.3"; so sánh phiên bản để lại tài liệu ẩn tới khi thoát Word; cửa sổ Word thu nhỏ thì không chèn được nhận xét.
- PowerPoint: lần chèn biểu đồ đầu trên bài mới để lại một bố cục thừa; không có API ghi chú diễn giả, số liệu biểu đồ.
- OpenRouter: không thu hồi được khóa, mỗi lần đăng nhập cấp thêm một khóa; sau Đăng xuất rồi đăng nhập lại, nguồn đang chọn không tự quay về OpenRouter.
- Nguồn có danh sách model cố định (MiniMax, Kimi…): key chỉ được kiểm ở lượt gọi đầu.

### Chờ anh quyết hoặc làm

- Tối ưu theo mã add-in chính thức (25/11, báo cáo mục 7): mức 1, 2, 3 xong (mức 3 ở 4.17.0). Bản trích gốc lời nhắc chỉ ở `_luu-cuc-bo/rea-claude-office-2026-11-25/` (ngoài Git, không đưa lên kho công khai).

- Rà soát kiến trúc 20/11, bảng "Phát hiện để lại" (`tasks/viec-le/ra-soat-kien-truc/ra-soat-kien-truc-2026-11-20.md`): nhóm 1 (B1, B2, B3, B4, B6, B23) xong ở 4.11.13. Nhóm 2 xong ở 4.15.0 (B8 không làm). Còn: nhóm 3 xong ở 4.13.3 trừ tách công cụ theo ứng dụng (B10 phần 2); nhóm 4 bảo mật nhỏ (B11, B12, B17, B19, B20, B21, B22; B5 xong ở 4.15.0). B14 giữ nguyên.
- Hai thư mục mồ côi `dot-0a-office-addon-compare-47836b`, `dot-4-a4b5e7` trong `.claude\worktrees\`: đưa vào Thùng rác.
- Xóa các khóa "AI Da Nguon (Office)" thừa ở `openrouter.ai/settings/keys`.

## 5. Bước tiếp theo

0. Mức 3 xong (4.17.0): PowerPoint storyboard + văn phong (việc 10), biểu đồ theo đặc tả (việc 11). Để lại: chưa Chạy thật với model xem AI có ghi key_message và sửa theo style_warnings không (Miễn phí làm được, không tốn tiền); bỏ `sort` của biểu đồ (AI tự sắp số liệu). Một lần `npm test` trượt 1 ca ngay sau `git commit -am` (27/11), chạy lại 5 lần đều đạt, chưa rõ ca nào — gặp lại thì ghi tên ca. Chọn tiếp theo mục 1. Từ 27/11 anh giao tự triển khai và tự quyết phát hành; chỉ hỏi khi đổi kiến trúc, đụng tiền, mật khẩu, lộ thông tin.
1. Chọn tiếp: phần để sau của nhóm 2–3. Còn để sau: khối `onEvent` của `Chat.tsx`, tách công cụ theo ứng dụng (B10 phần 2). 4.13.3 đã phát hành, Tự kiểm tra đạt. 4.13.2 xong (phát hành, Tự kiểm tra, Chạy thật đạt). 4.13.1 xong (Tự kiểm tra đạt; Chạy thật Trả phí 26/11, Anthropic · Claude Sonnet 5.5: AI tự soát trước khi ghi, căn cứ chưa kiểm được nói rõ; 6/6 tiêu chí — `tasks/viec-le/rea-claude-chinh-thuc/` mục 9). Chọn tiếp: mức 2 của báo cáo, hoặc nhóm 3 rà soát kiến trúc (khung mở nhanh hơn), hoặc lập kế hoạch nhóm 2.
2. **Buổi kiểm bù** khi anh đã nạp tiền hoặc có key: chạy hết nhóm "chưa kiểm được" ở mục 4.
3. Việc tùy chọn (mục D của `tasks/dot/dot-5/dot-5.md`), anh chọn: tải bản ghi hội thoại, gợi ý gạch chân bấm để chấp nhận trong Word, connectors (MCP), công cụ chạy mã Office.js do AI viết, phím tắt mở khung, đọc một trang web cụ thể, Outlook, Sign in with Poe, hàm `=AI(…)` trong ô Excel. Đã làm: điền mẫu (Đợt 8), bộ nhớ dài hạn (Đợt 12 E).

Câu mở đầu phiên sau: `Đọc tasks/HANDOFF.md mục Bước tiếp theo rồi làm tiếp.`

## 6. Bài học mới của phiên

- (25/11/2026) Gói `pivot.claude.ai/m-addin` (9,8 MB, một tệp) làm REA `analyze_javascript_application` chạm trần trạng thái bí danh; `recover_javascript_sources` cần Linux. Cách chạy được: phân tích cú pháp bằng `acorn` (không chạy mã) và trích chuỗi dài — lời nhắc hệ thống từng ứng dụng nằm nguyên văn trong gói. Script ở `_luu-cuc-bo/rea-claude-office-2026-11-25/strings.cjs`.

- (24/11/2026) Tệp bàn giao lại phình tới 87 KB sau lần viết gọn 18/11, vì "giữ nguyên văn mục Quyết định và Bài học" cộng với mỗi phiên chỉ thêm. Nay Bài học có tệp riêng (`docs/bai-hoc.md`), tệp này có trần 15 KB ghi ở đầu và ở CLAUDE.md.
- (24/11/2026) Một nhánh sửa lỗi thật (rà soát kiến trúc 20/11, 8 chỗ sửa + 125 dòng kiểm thử) nằm quên 4 ngày ở worktree vì HANDOFF ghi "chưa gộp" lẫn trong đoạn dài. Cuối mỗi phiên: `git branch --no-merged main` và liệt kê mọi nhánh chưa gộp ở mục "Chờ anh quyết".
