# Bàn giao – Add-in AI đa nguồn cho Office

**Cập nhật:** 15/11/2026 · Thư mục: `D:\ADD IN` · Mã nguồn: https://github.com/DungHoang1985/ai-da-nguon (công khai, nhánh `main`) · Bản chạy: https://dunghoang1985.github.io/ai-da-nguon/ (nhánh `gh-pages`) · **Phiên bản đang chạy: v3.5** (hiện ở đầu khung add-in)

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
- **Đang làm:** Đợt 0A (bắt đầu 15/11/2026, phiên Claude Code). Bước 1 xong: `D:\ADD IN` là bản làm việc của `main`; mã bản 0 cũ chuyển vào `_ban-0-cu/` (không đưa lên kho, chờ xóa). Tiến độ từng bước: xem `git log`.
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
  - Lỗi đổi nhà cung cấp chưa tái hiện trong Office (bước 4.1 của `dot-0.md`).
  - Phải thử trong Office trước khi cam kết: bảng ở mục 8 của `ke-hoach-v4.md` (đọc số điều khoản tự đánh, tải tệp từ khung add-in, đăng nhập OpenRouter trong hộp thoại, ba ứng dụng dùng chung kho…).
  - Vị trí nút chọn model, lịch sử trong khung Claude chính thức chưa có nguồn: chụp add-in chính thức đang cài trên máy trước khi dựng bố cục.
  - Chưa chạy thật: tìm trên web, mở lại/xóa hội thoại cũ, tự tóm tắt hội thoại dài, bấm địa chỉ ô/đoạn/slide, nhập kỹ năng qua giao diện, ảnh đính kèm, lưu "Chỉ dẫn riêng" → kiểm nốt ở Đợt 0.
  - Chưa chạy thật với OpenRouter/Gemini/Groq → nút "Kiểm tra nhà cung cấp" ở Đợt 0.
  - Ba khoảng trống lớn nhất với việc pháp lý – BĐS: Word chưa đọc số điều khoản tự đánh; PDF quét chưa đọc được; chưa ghi chéo ứng dụng.
  - Lỗi đã thấy khi đọc mã, chưa sửa: dừng/lỗi giữa chừng thì AI quên việc vừa làm; không tự thử lại khi 429/5xx; ngưỡng tóm tắt cố định; chưa dùng bộ nhớ đệm lời nhắc của Anthropic (tốn tiền).
  - Ba mục phải thử kỹ thuật trước khi cam kết: đăng nhập OpenRouter trong hộp thoại Office; khung ứng dụng này gửi lệnh cho khung ứng dụng kia; connectors qua MCP của Claude API.
  - Presentation1 (bài nháp chưa lưu) có 2 bố cục thừa — đóng không lưu. Tệp thử `kiem-thu\tep-thu-AIDN.pdf|.docx` — giữ để kiểm đính kèm.
- **Bước tiếp theo:** dự án **chuyển sang làm trong Claude Code** (anh quyết định 14/11/2026), mở tại thư mục `D:\ADD IN`. Tệp bàn giao: `CLAUDE.md` (Claude Code tự đọc), `tasks\dot-0.md` (việc cụ thể của Đợt 0). Làm **Đợt 0A** trước. Câu mở đầu: `Làm Đợt 0A theo tasks/dot-0.md, bắt đầu từ bước 1.` (Opus 5.5; model Fable hết hạn mức tới 19h ngày 20/11/2026)
  - Chưa biết Claude Code trên máy anh có điều khiển được màn hình Office không. Nếu không: các mục phải kiểm trong Office ghi "Chưa kiểm được" và kiểm bù bằng một phiên Claude có điều khiển máy.
  - Đã kiểm 14/11: máy anh có Git 2.56, Node 24.14 và quyền đẩy lên GitHub (`git push --dry-run` đạt).
- **Bài học:**
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
