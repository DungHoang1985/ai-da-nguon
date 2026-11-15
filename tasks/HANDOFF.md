# Bàn giao – Add-in AI đa nguồn cho Office

**Cập nhật:** 15/11/2026 · Thư mục: `D:\ADD IN` · Mã nguồn: https://github.com/DungHoang1985/ai-da-nguon (công khai, nhánh `main`) · Bản chạy: https://dunghoang1985.github.io/ai-da-nguon/ (nhánh `gh-pages`) · **Phiên bản đang chạy: v3.6** (phát hành tự động 15/11/2026) (hiện ở đầu khung add-in)

> Từ 15/11/2026, `D:\ADD IN` là bản làm việc của nhánh `main` (đã đồng bộ ở bước 1 Đợt 0A); `docs/`, `tasks/`, `kiem-thu/` đã có trong kho.

- **Đã xong:**
  - Bản 1: nhập key (lưu cục bộ), tự liệt kê model theo 11 nhà cung cấp, chat streaming, chèn câu trả lời.
  - Bản 2 – chế độ tác tử (AI tự thao tác Excel/Word/PowerPoint, Hoàn tác, Track Changes, nút Tự kiểm tra).
  - **Bản 3 – theo add-in Claude chính thức + kỹ năng chính thức tích hợp sẵn:**
    - Excel: tìm kiếm, quét lỗi công thức, lần ô tham chiếu/phụ thuộc, lọc, PivotTable, định dạng có điều kiện, danh sách chọn, ghi chú ô, chèn/xóa hàng cột, đổi tên sheet, cố định dòng/cột; cảnh báo trước khi ghi đè/xóa.
    - Word: đọc/trả lời/đánh dấu đã xử lý nhận xét (kèm đoạn được neo), đọc và chấp nhận/từ chối sửa đổi được theo dõi, đọc/điền bảng có sẵn, chuyển tới đoạn.
    - PowerPoint: biểu đồ gốc sửa được dữ liệu (pptxgenjs → `insertSlidesFromBase64`, dùng bố cục "Chỉ tiêu đề" và màu chủ đề của mẫu), sơ đồ quy trình tự dàn trang, nhân bản/di chuyển/xóa slide, sửa/xóa hình; Hoàn tác cả khi slide được khôi phục bằng bản sao.
    - Chung: chỉ dẫn riêng theo ứng dụng; lịch sử hội thoại (IndexedDB); tự tóm tắt hội thoại dài; tệp đính kèm PDF/DOCX/TXT/CSV/MD/ảnh; lệnh "/"; nhập SKILL.md/.zip/.skill/gói plugin Claude; đọc chéo Excel–Word–PowerPoint; tìm trên web (OpenRouter/Anthropic/Gemini); địa chỉ ô/đoạn/slide bấm được; số phiên bản ở đầu khung.
    - 17 kỹ năng tích hợp sẵn, chuyển thể tiếng Việt từ skill chính thức `xlsx`, `docx`, `pptx`, `dataviz` và plugin `legal`.
  - Kiểm chứng trên máy anh (bản 3.4–3.5): Tự kiểm tra Excel 22/22, Word 14/14, PowerPoint 14/14. Chạy thật với key Anthropic (Claude Opus 5.5): PowerPoint tự áp dụng kỹ năng `thiet-ke-bieu-do`, chèn biểu đồ gốc, Hoàn tác 2 thao tác; đọc 2 tệp đính kèm PDF + DOCX; đọc chéo Excel đang mở. Menu "/" hiện đúng kỹ năng (Excel).
  - **Phiên 14/11/2026 – chỉ lập kế hoạch, không sửa mã:**
    - Audit mã nguồn trên bản sao kho (`main`, commit 14/11 00:03): `npm test` 71/71, kiểm kiểu không lỗi, đóng gói đạt. Kết luận: nền tốt, không viết lại; 9 vấn đề cần xử lý và 6 phần thừa — chi tiết ở `tasks\ke-hoach-v4.md` mục A.
    - Rà lại nguồn chính thức (9 trang claude.com/docs/office-agents + 4 bài hỗ trợ) và viết lại `docs\so-sanh-claude-chinh-thuc.md` thành một bảng 55 dòng: Có 29 · Một phần 11 · Chưa 11 · Không làm được 1 · Không cần/hai bên đều không 3.
    - So sánh 3 hướng đăng nhập (mục B của kế hoạch): OpenRouter một lần bấm — được phép; gói Claude qua Agent SDK — vùng xám, không khuyến nghị; tín dụng API kèm gói — được phép, chỉ Max/Ultra/Team.
    - Lập kế hoạch 5 đợt + 1 đợt tùy chọn, mỗi đợt có tiêu chí nghiệm thu và cách tự kiểm (mục C của kế hoạch).
  - **Phiên 14–15/11/2026 (Claude Code) – rà lại kế hoạch, không sửa mã:**
    - Lập kế hoạch bản 2, **anh duyệt 15/11**: nay là `tasks\ke-hoach-v4.md` (bản cũ lưu ở `tasks\ra-soat-2026-11-15\ke-hoach-v4-ban-1.md`); `tasks\dot-0.md` đã viết lại cho Đợt 0A. Bằng chứng: `tasks\ra-soat-2026-11-15\` (đọc `DOC-TRUOC.md`).
    - Kiểm trên máy anh: `npm test` 71/71, kiểm kiểu sạch, đóng gói đạt; Git 2.56, Node 24.14; `git push --dry-run` thành công (đã có quyền đẩy, không cần đăng nhập lại); Office 16.0.20527 Current Channel, WebView2 157.
    - Anh nêu thêm 4 việc: (1) Enter để gửi (hiện chỉ Ctrl+Enter, `Chat.tsx:974–976`); (2) phần cài đặt ban đầu ẩn vào Cài đặt để vùng chat rộng; (3) lỗi mất kết nối khi đổi nhà cung cấp; (4) rà chức năng gốc của Word, Excel, PowerPoint để đề xuất tính năng.
    - Lỗi đổi nhà cung cấp: khoanh được 4 nguyên nhân (SambaNova và NVIDIA chặn gọi từ add-in; khung giữ danh sách và model của nguồn cũ khi tải lỗi; model tự nhảy về mục đầu theo chữ cái; mọi `TypeError` đều hiện "Không kết nối được"). Chưa tái hiện trong Office.
    - Phát hiện mới mức cao: Office lưu dữ liệu add-in tách theo tài khoản Office; máy anh có 2 kho (`Wef\webview2`: một `_ADAL`, một `_LiveId`) → cần Sao lưu/Khôi phục. Mã Word có 20 phát hiện (12 mức cao). Bảng so sánh sai 3 chỗ, thiếu khoảng 12 khả năng mới của bản chính thức.
    - Chưa rà: mã lõi tác tử (tôi chỉ đọc trực tiếp `loop.ts`, `llm.ts`, `compaction.ts`, `errors.ts`), mã Excel, PowerPoint, kỹ năng, tệp đính kèm; tra cứu nền tảng Office và API nhà cung cấp. Lý do: 10/14 tác tử bị dừng vì hết hạn mức model.
- **Đợt 0A (phiên 15/11/2026, Claude Code) — mã xong bước 1–8 và nâng v3.6, đã phát hành; kiểm trong Office mới làm một phần:**
  - Bước 1 (6e4d9b7): `D:\ADD IN` là bản làm việc của `main`; `docs/`, `tasks/`, `kiem-thu/` đã vào kho; `_ban-0-cu/` đã chuyển vào Thùng rác (anh duyệt).
  - Bước 2 (6a828ed): phát hành tự động `.github/workflows/phat-hanh.yml`. Bằng chứng: run 43657987624 xanh, gh-pages có commit của bot, Pages đổi theo; nhánh thử trượt kiểm thử → bước phát hành bị bỏ qua (run 43658148652).
  - Bước 3 (9532f7a): gỡ `chat.ts`, gói đệm IE, 10 thư viện; số phiên bản lấy từ `package.json`; `image-size` ép 2.1.0 (npm audit bản chạy: 0).
  - Bước 4 (f7d941b, bd0cfb4): sửa lỗi đổi nhà cung cấp + nút Kiểm tra nhà cung cấp. **Tái hiện trong Word (v3.5):** OpenRouter bật "Miễn phí" → đổi sang Anthropic: 0/15 model, gửi model OpenRouter cho Anthropic. **Nguyên nhân "mất kết nối":** OpenAI (key dạng sk-… sai/hết hạn) và Cerebras trả trang lỗi thiếu CORS → bản cũ báo nhầm "Không kết nối được". **Chạy thật trong Word (15/11):** từng nguồn 5/8 đạt (Gemini, DeepSeek: tài khoản hết tiền; OpenAI: lỗi bị giấu — đã sửa cách báo, chưa chạy lại); đổi nguồn đạt 20/20 cặp giữa 5 nguồn chạy được (OpenRouter, Groq, Anthropic, Cerebras, Cohere). Bằng chứng: `tasks/dot-0a-bang-chung/`.
  - Bước 5 (0385c02) Enter để gửi · Bước 6 (c48b20f) bố cục một dòng đầu, bỏ công tắc Track Changes · Bước 7 (6ada3e8) sao lưu/khôi phục · Bước 8 (2831125) ba lỗi Word · 69627fb nâng 3.6.
  - Kiểm thử máy: `npm test` 118/118 (71 cũ − 5 của chat.ts + 52 mới), tsc sạch. Bố cục mới đo trong trình duyệt (330×620): vùng hội thoại 71,7%.
  - **Bước 9 (phiên 15/11 buổi trưa, trên máy, có điều khiển Office) — kiểm trong Office bản 3.6 xong; bảng kết quả đủ ở `tasks/dot-0b.md`, ảnh ở `tasks/dot-0a-bang-chung/08`–`20`:**
    - Đạt: đầu khung v3.6, vùng hội thoại khoảng 79% chiều cao khung; Enter gửi / Shift+Enter xuống dòng; gõ ba câu có dấu bằng Unikey không bị gửi giữa chừng; đổi qua lại Anthropic ↔ Groq mỗi nguồn về đúng model đã dùng; OpenAI báo đúng loại nguyên nhân; sao lưu → xóa dữ liệu thử → khôi phục gộp (key còn nguyên); tải tệp từ khung add-in ra thư mục Tải xuống; Excel 22/22; Track Changes đang tắt mà AI chèn và thay chữ vẫn ra sửa đổi có theo dõi (chạy thật, Groq).
    - Không đạt / chưa kiểm được: Tự kiểm tra Word 16/17 (bước kiểm đo sai, đã sửa mã, chưa phát hành); PowerPoint lần đầu trên bài mới 13/14, lần hai 14/14; bộ gõ Telex của Windows chưa cài trên máy nên chưa kiểm.
    - Kiểm tra nhà cung cấp chạy lại 12:10: đạt 4/8 (Groq, Anthropic, Cerebras, Cohere); Gemini, DeepSeek hết tiền; OpenAI lỗi bị giấu; OpenRouter trượt vì model đang nhớ bị giới hạn nơi dùng.
    - Bảng so sánh sửa theo mục 7: 67 dòng (Có 25 · Một phần 17 · Chưa 21 · Không làm được 1 · Không cần/hai bên đều không 3). Tài liệu chính thức và `shortcuts.json` tải lại từ máy cùng ngày.
    - Mã đã sửa trên nhánh `claude/dot-0a-viec-con-lai`, **chưa commit, chưa đẩy, chưa phát hành**: bước Tự kiểm tra Word nói trên; `word_replace_text` hết lỗi khi model truyền `max_replacements` âm. `npm test` 120/120, tsc sạch.
- **Đang làm:** chờ anh cho phát hành hai sửa đổi của bước 9, rồi chạy lại Tự kiểm tra Word (phải ra 17/17). Sau đó vào Đợt 0B theo `tasks/dot-0b.md`.
- **Quyết định và lý do:**
  - AI sửa thẳng vào tài liệu + Hoàn tác; Word luôn ghi Track Changes; thao tác rủi ro luôn hỏi kèm lý do, "Cho phép hết lần này" không áp dụng cho thao tác rủi ro. Lý do: tự động tối đa nhưng an toàn như bản chính thức.
  - Kỹ năng chính thức được chuyển thể, không chép nguyên văn; không ghi số hiệu văn bản pháp luật trong kỹ năng (có kiểm thử chặn).
  - Tự chọn kỹ năng bằng từ khóa tiếng Việt không dấu + danh mục kỹ năng trong lời nhắc để model tự gọi `use_skill`. Audit 14/11: giữ nguyên.
  - Connectors: trước ghi "không làm". Cập nhật 14/11: xếp vào Đợt 5 tùy chọn, có thể làm một phần với key Anthropic, phải thử kỹ thuật trước.
  - **Duyệt kế hoạch v4 (14/11/2026): ĐÃ DUYỆT**, anh chọn:
    1. Thứ tự: Đợt 0 → 1 → 2 → 3 → 4 (Word pháp lý trước đăng nhập, vì giá trị cao nhất).
    2. Đăng nhập: OpenRouter một lần bấm + tín dụng API kèm gói. **Không làm** hướng gói Claude qua Agent SDK (vùng xám).
    3. Anh đang dùng gói **Max** → đủ điều kiện tín dụng API 100 USD/tháng (đối chiếu lại bài 17154008 trước khi làm Đợt 2).
    4. Dọn nguồn, làm cả ba: thay mã cũ trong `D:\ADD IN` bằng bản hiện hành; đưa `docs/`, `tasks/` lên kho GitHub công khai; bật phát hành tự động. Lưu ý khi làm: xóa tệp trên máy phải xin quyền xóa; rà lại `docs/`, `tasks/` không chứa thông tin riêng tư trước khi đẩy lên kho công khai.
  - **Duyệt kế hoạch bản 2 (15/11/2026): ĐÃ DUYỆT**, anh chọn:
    1. Thứ tự: 0A → 0B → 1A → 1B → 2 → 3 → 4 (Đợt 0 tách đôi; Word vẫn trước đăng nhập). Lý do tách: anh thấy bốn việc mình nêu được xử lý ngay ở bản 3.6.
    2. Bố cục khung chat kiểu **một dòng đầu** (bỏ hai thẻ Chat / Cài đặt, nhãn model cạnh ô nhập, Cài đặt mở bằng nút bánh răng), giống add-in Claude chính thức.
    3. Lỗi mất kết nối khi đổi nhà cung cấp: anh không nhớ rõ cặp nguồn → Claude tự thử mọi cặp nguồn anh có key trong Office.
    4. **Bỏ công tắc tắt Track Changes**: sửa đổi của AI trong Word luôn ghi có theo dõi.
- **Vấn đề mở:**
  - Tìm ra khi kiểm bản 3.6 (15/11), chưa sửa: (1) PowerPoint — lần chèn biểu đồ đầu tiên trên bài mới tinh làm mẫu thêm một bố cục (11 → 12), các lần sau dùng lại; (2) OpenRouter — model đang nhớ `thinkingmachines/inkling-2-small:free` chỉ chạy trong "agentic harnesses", add-in báo nhầm là "Key không đúng hoặc không có quyền"; (3) bấm Gửi khi bảng Lịch sử đang mở thì bảng vẫn che hội thoại; (4) Enter với bộ gõ Telex của Windows chưa kiểm (máy chưa cài bộ gõ này).
  - Tệp sao lưu thử `AI-Da-Nguon-sao-luu-20261115-1212-BI-MAT.json` còn trong thư mục Tải xuống của máy (chứa key thật): anh quyết giữ hay xóa.
  - Lỗi đổi nhà cung cấp: đã tái hiện và sửa ở bước 4 Đợt 0A (mục này của bản trước đã cũ).
  - Phải thử trong Office trước khi cam kết: bảng ở mục 8 của `ke-hoach-v4.md` (đọc số điều khoản tự đánh, tải tệp từ khung add-in, đăng nhập OpenRouter trong hộp thoại, ba ứng dụng dùng chung kho…).
  - Vị trí nút chọn model, lịch sử trong khung Claude chính thức chưa có nguồn: chụp add-in chính thức đang cài trên máy trước khi dựng bố cục.
  - Chưa chạy thật: tìm trên web, mở lại/xóa hội thoại cũ, tự tóm tắt hội thoại dài, bấm địa chỉ ô/đoạn/slide, nhập kỹ năng qua giao diện, ảnh đính kèm, lưu "Chỉ dẫn riêng" → kiểm nốt ở Đợt 0.
  - Chưa chạy thật với OpenRouter/Gemini/Groq → nút "Kiểm tra nhà cung cấp" ở Đợt 0.
  - Ba khoảng trống lớn nhất với việc pháp lý – BĐS: Word chưa đọc số điều khoản tự đánh; PDF quét chưa đọc được; chưa ghi chéo ứng dụng.
  - Lỗi đã thấy khi đọc mã, chưa sửa: dừng/lỗi giữa chừng thì AI quên việc vừa làm; không tự thử lại khi 429/5xx; ngưỡng tóm tắt cố định; chưa dùng bộ nhớ đệm lời nhắc của Anthropic (tốn tiền).
  - Ba mục phải thử kỹ thuật trước khi cam kết: đăng nhập OpenRouter trong hộp thoại Office; khung ứng dụng này gửi lệnh cho khung ứng dụng kia; connectors qua MCP của Claude API.
  - Presentation1 (bài nháp chưa lưu) có 2 bố cục thừa — đóng không lưu. Tệp thử `kiem-thu\tep-thu-AIDN.pdf|.docx` — giữ để kiểm đính kèm.
- **Bước tiếp theo (cập nhật 15/11 buổi trưa):** (1) anh cho phép thì commit và đẩy nhánh `claude/dot-0a-viec-con-lai` lên `main` (đẩy là phát hành), đặt con trỏ trong khung add-in bấm Ctrl+F5, chạy lại Tự kiểm tra Word — phải ra 17/17; (2) vào **Đợt 0B** theo `tasks/dot-0b.md`, trước đó rà nốt `loop.ts`, `llm.ts`, `compaction.ts`. Câu mở đầu phiên sau: `Làm Đợt 0B theo tasks/dot-0b.md.` — phải là phiên **trên máy**, không phải phiên đám mây.
- **Bước tiếp theo (bản cũ, để tham khảo):** dự án **chuyển sang làm trong Claude Code** (anh quyết định 14/11/2026), mở tại thư mục `D:\ADD IN`. Tệp bàn giao: `CLAUDE.md` (Claude Code tự đọc), `tasks\dot-0.md` (việc cụ thể của Đợt 0). Làm **Đợt 0A** trước. Câu mở đầu: `Làm Đợt 0A theo tasks/dot-0.md, bắt đầu từ bước 1.` (Opus 5.5; model Fable hết hạn mức tới 19h ngày 20/11/2026)
  - Chưa biết Claude Code trên máy anh có điều khiển được màn hình Office không. Nếu không: các mục phải kiểm trong Office ghi "Chưa kiểm được" và kiểm bù bằng một phiên Claude có điều khiển máy.
  - Đã kiểm 14/11: máy anh có Git 2.56, Node 24.14 và quyền đẩy lên GitHub (`git push --dry-run` đạt).
- **Bài học:**
  - **Phiên đám mây không kiểm được Office và không gọi được nhà cung cấp** (không có Office, không có key, một số địa chỉ như `pivot.claude.ai` bị chặn). Việc cần Office — Tự kiểm tra, Chạy thật, Kiểm tra nhà cung cấp, chụp khung add-in — phải chạy trong phiên trên máy có điều khiển máy tính. Phiên đám mây chỉ nên làm: đọc mã, kiểm thử máy, đối chiếu tài liệu; kết luận của phiên đám mây về tài liệu vẫn phải soát lại (lần này phiên đám mây báo "không thấy phát hiện va chạm", thực ra có ở trang OpenTelemetry).
  - Phiên Claude Code mở trong worktree có thể được tạo từ nhánh mặc định `gh-pages` (chỉ có bản đóng gói, không có `tasks/`, `src/`). Thấy thiếu tệp thì tạo nhánh mới từ `main` (`git switch -c <tên> main`), không `reset --hard`.
  - Tự kiểm tra "đạt" chưa chắc đã đo đúng: bước nào chỉ so số đếm của công cụ trả về (ví dụ `replaced === 1`) mà không đọc lại tài liệu thì có thể đạt giả; bước đếm Track Changes phải làm trên chữ thường, vì sửa trong đoạn đang là sửa đổi chờ duyệt thì Word gộp, số đếm không tăng. Bước Tự kiểm tra mới phải chạy trong Office ít nhất một lần trước khi ghi "đạt".
  - PowerPoint: muốn biết bước "Biểu đồ gốc" có thêm bố cục không thì phải chạy trên bài **mới tinh**; chạy lại trên bài đã thử sẽ đạt giả vì bố cục thừa đã có sẵn.
  - Model yếu hay truyền tham số số học kiểu `-1` = "tất cả": mọi tham số giới hạn của công cụ phải kẹp về số nguyên dương.
  - Điều khiển máy: lệnh gõ chữ thường không đi qua Unikey, còn bấm từng phím thì có → muốn thử gõ tiếng Việt phải bấm từng phím Telex. Khi khung add-in đang giữ con trỏ, gõ vào thân tài liệu Word không ăn → nhờ AI chèn nội dung thử thay vì gõ tay. Mở Excel, PowerPoint bằng nút Add-ins → Developer Add-ins → AI Đa Nguồn nếu chưa thấy nút trên thẻ Home.
  - Word, Excel, PowerPoint trên máy anh (cùng tài khoản Office) dùng chung một kho dữ liệu add-in: model vừa chọn ở Word hiện đúng ở Excel và PowerPoint (có ích cho Đợt 3).
  - Lỗi "không kết nối được" trong add-in chưa chắc là mạng: nhà cung cấp trả trang lỗi thiếu tiêu đề CORS thì trình duyệt giấu lỗi thật. Đo bằng curl với Origin của add-in, và trong add-in gọi thử chế độ no-cors để phân biệt (`fetchDiagnosed`).
  - Điều khiển máy: `open_application` Word mỗi lần lại mở thêm một tài liệu trống (Document2, 3…) → hạn chế gọi; Windows hay giữ màn hình nền ở lớp trên, cần quyền "File Explorer" (chỉ bấm) để bấm vào cửa sổ Word. Nội dung khung add-in cần quyền riêng cho `msedgewebview2.exe`, và chỉ xin được khi đang có khung add-in mở.
  - Quy trình phát hành: đẩy lên `main` là phát hành; muốn thử quy trình thì dùng nhánh `thu/…` (không đẩy gh-pages).
  - Lệnh xóa thư mục bị bộ phân loại chế độ tự động chặn → chuyển vào Thùng rác bằng PowerShell sau khi anh duyệt.
  - Phiên Claude Code có hạn mức sử dụng cứng theo model: phóng 14 tác tử song song làm phiên bị cắt hai lần, mất trắng lượt đầu. Chia đợt tối đa 4 tác tử, việc quan trọng chạy trước; tác tử tra cứu dùng model nhẹ; bị ngắt thì đọc `journal.jsonl` của workflow trước khi chạy lại.
  - Vùng tạm (scratchpad) của phiên bị xóa và đổi đường dẫn khi phiên khởi động lại: kết quả cần giữ phải chép ngay vào `tasks\`.
  - Dữ liệu add-in (key, lịch sử, kỹ năng) nằm trong `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2\<tài khoản Office>`: **không xóa thư mục `Wef`**; đổi tài khoản Office là sang kho khác.
  - Bốn bài hỗ trợ của Anthropic về add-in (12650343, 13521390, 14465370, 13892150) đã gộp về `claude.com/docs/office-agents/`; tài liệu chính thức nay có 14 trang (thêm performance, outlook, opentelemetry).
  - **Trước khi audit hay sửa: kiểm xem thư mục trên máy có đúng là mã đang chạy không** (so Id trong manifest, số phiên bản). Phiên 14/11 suýt audit nhầm bộ khung bản 0.
  - Trang claude.com/docs/office-agents không tải trực tiếp được; tìm kiếm trước rồi tải bản `.md` của từng trang (`…/excel.md`, `…/word.md`…).
  - Bài hỗ trợ 15036540 đã đổi nội dung nhiều lần trong 2026 (tín dụng Agent SDK bị hủy, thay bằng tín dụng API hằng tháng ở bài 17154008) → mọi kết luận về đăng nhập và chi phí phải đối chiếu lại ngay trước khi làm Đợt 2.
  - `package-lock.json` / `.npmrc` của mẫu OfficeDev trỏ registry riêng của Microsoft (403) → dùng `registry.npmjs.org`.
  - Kiểm CORS và Office.js phải chạy trong Office thật; dùng nút Tự kiểm tra để có bằng chứng.
  - Anh không muốn gõ lệnh/cài tay: mọi bước phải đóng gói thành 1 tệp bấm đúp hoặc Claude tự làm. Anh đánh giá "đạt" khi AI thao tác thật trên tài liệu, không chỉ chat.
  - Khóa registry WEF\Developer: tên giá trị phải là Id trong manifest (GUID).
  - Excel `getCellProperties` trả về cả trường không ghi lại được (address, viền trong) → lọc trước khi `setCellProperties`.
  - Word: "ListBullet"/"ListNumber" không phải tên kiểu dựng sẵn hợp lệ → dùng `startNewList()`; văn bản trống: nội dung chèn "cuối" nằm trước đoạn trống ban đầu.
  - Điều khiển máy: cửa sổ Claude hay nổi lên che góc phải và chặn cú bấm ở nửa trên khung add-in → bấm ô nhập phía dưới rồi Shift+Tab/Tab (từ ô nhập: Shift+Tab ×12 → Phải → Enter = thẻ Cài đặt); bấm thanh tiêu đề hoặc vùng trống của Office để đưa Office lên trước rồi mới chụp.
  - Ctrl+F5 khi con trỏ đang trong khung add-in = tải lại bỏ qua bộ nhớ tạm. Nếu con trỏ đang ở Excel thì Ctrl+F5 thu nhỏ cửa sổ sổ tính.
  - Ô nhập trong WebView2 bị trình duyệt tự điền → đặt `autoComplete="off"`.
  - `webpack.ProvidePlugin({ Promise: "es6-promise" })` của mẫu OfficeDev thay Promise gốc → pdf.js lỗi "withResolvers is not a function" → bỏ.
  - pptxgenjs: biểu đồ 1 chuỗi nhiều màu ghi `<a:srgbClr val="accent1"/>` (sai chuẩn) → PowerPoint báo InvalidArgument; chỉ dùng 1 màu cho 1 chuỗi (có kiểm thử).
  - `insertSlidesFromBase64` + "Use destination theme": PowerPoint ghép bố cục theo kiểu (`sldLayout@type`), không theo tên → gắn kiểu cho bố cục.
  - Flex column trong khung add-in thấp: phần tử không đặt `flex-shrink: 0` (menu "/") bị co về 0.
  - Thư mục `/tmp/pages` (nhánh gh-pages) không có remote → đẩy bằng URL đầy đủ của kho.
