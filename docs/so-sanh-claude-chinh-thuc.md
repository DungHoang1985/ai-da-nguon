# So sánh AI Đa Nguồn (v3.7.0) với add-in Claude chính thức

Cập nhật 15/11/2026 (bước 9 Đợt 0A): sửa theo mục 7 của `tasks/ke-hoach-v4.md` sau khi đối chiếu lại 13 trang tài liệu chính thức (bản `.md`) và tệp `pivot.claude.ai/shortcuts.json` tải từ máy cùng ngày. Số dòng 1–55 giữ nguyên như bản 14/11; dòng 56–67 là khả năng mới của bản chính thức. Cập nhật 15/11/2026 buổi chiều (Đợt 0B, bản 3.7.0): sửa dòng 5, 9, 61, 66 và cột Bằng chứng của dòng 7, 8, 10, 13, 15, 22 sau khi kiểm bản 3.7.0 trong Office (bằng chứng: `tasks/dot-0b-bang-chung/`).

## Cách đọc cột "Bằng chứng"

- **Tự kiểm tra**: nút Tự kiểm tra chạy thao tác thật trong Office trên máy anh. Bản 3.7.0 (15/11 buổi chiều): Word 19/19 · Excel 22/22 · PowerPoint 13/14 lần đầu trên bài mới, 14/14 lần hai. Bản 3.6 (15/11 buổi trưa): Excel 22/22 (bản 3.6) · Word 17/17 (bản 3.6.1; bản 3.6 ra 16/17 vì một bước kiểm đo sai — xem dòng 42) · PowerPoint 14/14 ở lần chạy thứ hai (bản 3.6; lần đầu trên bài mới tinh 13/14 — xem dòng 53). Ảnh: `tasks/dot-0a-bang-chung/16`–`21`.
- **Chạy thật**: đã giao việc cho AI thật và thấy kết quả trong tài liệu (Claude Opus 5.5 với key Anthropic; ngày 15/11 thêm Groq · GPT OSS 20B).
- **Kiểm thử máy**: chỉ có kiểm thử tự động ngoài Office (`npm test`). Chưa thấy chạy trong Office.
- **Giao diện**: thấy hiện đúng trên màn hình, chưa thao tác thật.
- **—**: chưa có gì để kiểm.

Lưu ý chung: phần lớn mục "Chạy thật" chạy với key Anthropic. Nút Kiểm tra nhà cung cấp (bản 3.6, 15/11/2026, chuỗi "gọi công cụ → nhận kết quả → trả lời"): **đạt** Groq, Anthropic, Cerebras, Cohere; **không đạt** Gemini và DeepSeek (tài khoản hết tiền), OpenAI (lỗi bị giấu chi tiết, nay báo đúng loại nguyên nhân), OpenRouter (model đang nhớ bị OpenRouter giới hạn; buổi sáng cùng ngày đạt với model khác). Chi tiết: `tasks/dot-0a-bang-chung/08-kiem-tra-nguon-v3.6.txt`.

Nguồn (viết tắt ở cột 2): [X] Excel · [W] Word · [P] PowerPoint · [O] Outlook · [A] Work across apps · [C] Connectors and Skills · [D] Data storage · [Dt] Dictation · [3P] Third-party platforms · [Pf] Performance · [OT] OpenTelemetry · [SJ] tệp shortcuts.json của add-in. Đường dẫn ở cuối tệp. **"Không xác minh được"** = chi tiết từng có ở bốn bài hỗ trợ cũ (đã gộp về trang tài liệu), nay không còn trong tài liệu hiện hành.

## Bảng so sánh

| # | Khả năng | Bản chính thức | AI Đa Nguồn | Bằng chứng | Mức quan trọng (pháp lý – BĐS) |
|---|---|---|---|---|---|
| | **CHUNG** | | | | |
| 1 | Đăng nhập bằng tài khoản Claude, dùng hạn mức của gói | Có [X][W][P] | **Không làm được** (chỉ ứng dụng của Anthropic được dùng đăng nhập này; xem kế hoạch, phần đăng nhập) | — | Cao |
| 2 | Kết nối bằng token/key, không cần tài khoản Claude | Có, qua "Cloud provider or gateway" [3P] | **Có** — nhập key của 11 nhà cung cấp | Chạy thật: Kiểm tra nhà cung cấp 15/11 đạt 4 nguồn (Groq, Anthropic, Cerebras, Cohere) | Cao |
| 3 | Chọn model | Có, danh sách Claude chọn lọc [X] | **Có**, rộng hơn (cả model miễn phí); nhớ model theo từng nguồn | Chạy thật (Anthropic, Groq): đổi qua lại hai nguồn, mỗi nguồn về đúng model đã dùng | Cao |
| 4 | AI tự đọc và sửa tài liệu qua nhiều bước | Có | **Có** (54 công cụ) | Tự kiểm tra + Chạy thật | Cao |
| 5 | Hỏi xác nhận trước thao tác rủi ro; cảnh báo trước khi ghi đè | Có [X][W] | **Có** từ bản 3.7 — Excel, PowerPoint: có. Word: hỏi khi xóa đoạn, chấp nhận/từ chối tất cả sửa đổi, ghi đè ô bảng đang có dữ liệu, thay vùng chọn lớn (từ 3 đoạn hoặc 1.000 ký tự), thay quá 5 chỗ một lúc | Tự kiểm tra (Excel, PowerPoint). Word, ba cảnh báo mới: Tự kiểm tra bản 3.7.0 (hai bước mới đạt) + Chạy thật 15/11 (ghi đè ô bảng: khung hiện cảnh báo kèm giá trị cũ và mới trước khi ghi) | Cao |
| 6 | Ba mức quyền: Ask before edits / Accept all edits / Skip all permissions | Có [3P] | **Một phần** — hai mức cạnh ô nhập ("Hỏi trước" / "Tự sửa, hỏi khi rủi ro") và "Cho phép hết lần này"; không làm mức bỏ qua mọi xin phép (trái quy tắc dự án) | Chạy thật (thẻ duyệt Cho phép / Cho phép hết lần này / Bỏ qua) | Trung bình |
| 7 | Chỉ dẫn riêng lưu sẵn cho từng ứng dụng | Có [X][W][P] | **Có** | Chạy thật 15/11: lưu, xóa, khôi phục từ tệp sao lưu (Word); AI làm theo chỉ dẫn vừa lưu (bản 3.7.0) | Cao |
| 8 | Lịch sử hội thoại lưu trên máy, xóa được | Có, 50 hội thoại gần nhất cho mỗi người dùng, mỗi tổ chức, mỗi ứng dụng [D] | **Có**, 60 hội thoại; thêm Sao lưu/Khôi phục ra tệp | Chạy thật 15/11: xóa một hội thoại, khôi phục lại từ tệp sao lưu; mở lại hội thoại cũ rồi hỏi tiếp, AI nhớ lượt trước (bản 3.7.0) | Trung bình |
| 9 | Tự thu gọn hội thoại dài | Có [X] | **Có** — từ bản 3.7 ngưỡng theo sức chứa của từng model (60% sức chứa; token thực đo vượt 75% cũng tóm tắt); giữa hai lần tóm tắt lịch sử không bị sửa | Chạy thật 15/11 (Word, Claude Haiku 5.5): tóm tắt ở lượt 11, sau đó AI vẫn nhớ đúng việc và số liệu lượt đầu | Trung bình |
| 10 | Tệp đính kèm làm nguồn | Có, kể cả ở chế độ third-party/gateway (ở chế độ đó tệp tải lên chỉ không dùng được cho chạy mã) [3P] | **Một phần** — PDF có lớp chữ, DOCX, TXT/CSV/MD/JSON, ảnh. **PDF quét (scan) chưa đọc được** | Chạy thật (PDF + DOCX); ảnh: Chạy thật 15/11 (bản 3.7.0, AI tả đúng ảnh) | **Cao** — hồ sơ pháp lý phần lớn là bản quét |
| 11 | Kỹ năng (Skills) tự áp dụng theo việc | Có [C] | **Có** — 17 kỹ năng dựng sẵn | Chạy thật (1 kỹ năng); còn lại Kiểm thử máy | Cao |
| 12 | Gọi kỹ năng bằng "/", lọc theo ứng dụng | Có [C] | **Có** | Giao diện (Excel) | Trung bình |
| 13 | Tải kỹ năng riêng lên, lưu trên máy | Có [D] | **Có** (SKILL.md, .zip, .skill, gói plugin) | Chạy thật 15/11 (bản 3.7.0): nhập một tệp SKILL.md, gọi bằng "/", xóa. Tệp .zip, .skill, gói plugin: Kiểm thử máy | Cao |
| 14 | Connectors (S&P, LSEG, Daloopa, MCP tùy biến) | Có, nút "+" [C] | **Chưa** — với key Anthropic có thể làm một phần, cần thử kỹ thuật | — | Thấp |
| 15 | Tìm kiếm web | Có [3P] | **Một phần** — theo nhà cung cấp (OpenRouter, Anthropic, Gemini) | Chạy thật 15/11 với Anthropic (bản 3.7.0); OpenRouter, Gemini: Kiểm thử máy | Cao (tra cứu văn bản) |
| 16 | Đọc nội dung một trang web cụ thể | Có [3P] | **Chưa** | — | Trung bình |
| 17 | Chạy mã để tính toán, phân tích | Có [3P] | **Chưa** | — | Thấp |
| 18 | Đọc giọng nói | Có, chỉ bản máy tính [Dt] | **Chưa** — dùng Win+H của Windows thay được | — | Thấp |
| 19 | Chéo ứng dụng: **đọc** tệp đang mở ở ứng dụng khác | Có [A] | **Có** — qua ảnh chụp nội dung, tối đa 30 phút tuổi | Chạy thật (PowerPoint đọc Excel) | Cao |
| 20 | Chéo ứng dụng: **ghi** sang tệp đang mở ở ứng dụng khác | Có [A] | **Chưa** | — | Cao (Excel → Word/PowerPoint báo cáo) |
| 21 | Hội thoại đi theo khi đổi ứng dụng; chỉ báo ứng dụng đang kết nối | Có [A] | **Chưa** | — | Trung bình |
| 22 | Trích dẫn bấm được (ô, đoạn, slide) | Có [X][W] | **Có** | Chạy thật 15/11 (bản 3.7.0): bấm [¶3] trong Word, Sheet1!D7 trong Excel, "slide 1" trong PowerPoint đều chuyển đúng chỗ | Cao |
| 23 | Outlook | Có (beta, đã có trang tài liệu đầy đủ): phân loại hộp thư, nháp trả lời theo giọng văn, tóm tắt chuỗi thư, đọc .docx/.xlsx đính kèm, tìm giờ họp; tham gia chéo ứng dụng [O] | **Chưa**, ngoài phạm vi (Đợt 5 nếu cần) | — | Thấp |
| 24 | Quản trị doanh nghiệp, Compliance API, nhật ký OpenTelemetry | Có | **Không cần** (dùng cá nhân) | — | Thấp |
| | **EXCEL** | | | | |
| 25 | Hỏi đáp nhiều sheet, giải thích công thức | Có [X] | **Có** | Tự kiểm tra | Trung bình |
| 26 | Đổi giả định, giữ nguyên quan hệ công thức | Có [X] | **Có** | Tự kiểm tra (ghi ô); chưa chạy thật trên mô hình | Trung bình |
| 27 | Tìm gốc lỗi #REF!, #DIV/0!, lần nguồn số liệu | Có [X] | **Có** (quét lỗi, lần ô trước/sau); tham chiếu vòng: chưa | Tự kiểm tra | Trung bình |
| 28 | Tạo mô hình, điền mẫu có sẵn | Có [X] | **Có** | Tự kiểm tra thao tác; chưa chạy thật | Trung bình |
| 29 | Sắp xếp, lọc, định dạng có điều kiện, danh sách chọn | Có [X] | **Có** | Tự kiểm tra | Trung bình |
| 30 | PivotTable | Có ("edit pivot tables"; tài liệu không tách tạo/sửa) [X] | **Một phần** — tạo mới; chưa sửa Pivot có sẵn | Tự kiểm tra | Thấp |
| 31 | Biểu đồ | Có (trang tổng quan nêu "charts"; không tách tạo/sửa) | **Một phần** — tạo mới; chưa sửa biểu đồ có sẵn | Tự kiểm tra | Trung bình |
| 32 | Tô sáng ô vừa sửa, kèm ghi chú giải thích | Không xác minh được trong tài liệu hiện hành (lần rà 14/11 còn thấy "tô sáng ô" ở trang sản phẩm; phần ghi chú giải thích hết nguồn) | **Một phần** — bấm vào bước để nhảy tới ô; không tô màu | Tự kiểm tra | Trung bình |
| 33 | Nhật ký phiên ghi vào một sheet ("Claude Log") | Có, chỉ Excel, **mặc định tắt**; bật thì giữ sheet "Claude Log" ở đầu, mỗi lượt trao đổi một dòng (ngày, tóm tắt yêu cầu, đã làm gì, kết quả); sheet đi theo tệp khi chia sẻ [D] | **Chưa** (Đợt 4) | — | Trung bình (dấu vết để soát lại) |
| 34 | Lấy bảng từ PDF, kể cả bản quét, đưa vào mẫu | Không xác minh được | **Một phần** — chỉ PDF có lớp chữ | Chạy thật (đọc PDF chữ) | Cao |
| 35 | Chuẩn bị in (đường lưới, vùng in) | Không xác minh được | **Chưa** (Đợt 4, giữ như lựa chọn riêng) | — | Thấp |
| 36 | Hỏi trước khi dùng hàm nguy hiểm (WEBSERVICE, INDIRECT…) | Không xác minh được | **Chưa** (Đợt 4, giữ như lựa chọn riêng) | — | Thấp |
| 37 | Hoàn tác | Không xác minh được cho Excel (trang Word nêu hoàn tác bằng Ctrl+Z chuẩn của Word [W]) | **Có, hơn** — nút Hoàn tác riêng | Tự kiểm tra + Chạy thật | Cao |
| 38 | Macro/VBA, Data table | Không hỗ trợ [X] | Không hỗ trợ | — | — |
| | **WORD** | | | | |
| 39 | Hỏi đáp về văn bản, trích dẫn mục bấm được | Có [W] | **Có** (dạng [¶ số đoạn]) | Tự kiểm tra (đọc); bấm: Kiểm thử máy | Cao |
| 40 | Hiểu đánh số pháp lý nhiều cấp, thuật ngữ định nghĩa, dẫn chiếu chéo | Có [W] | **Một phần** — **chưa đọc số điều khoản do Word tự đánh**, nên AI có thể không biết đoạn nào là "Điều 5.3" | — | **Cao nhất** |
| 41 | Sửa đoạn đang chọn, giữ kiểu chữ và đánh số | Có [W] | **Một phần** — sửa được vùng chọn và đoạn, nhưng nhiều kiểu sửa rơi về thay cả đoạn; việc giữ định dạng trong đoạn và đánh số chưa được kiểm (Đợt 1A) | Chưa có bước Tự kiểm tra cho sửa vùng chọn; sửa một đoạn "chỉ phần khác": Tự kiểm tra | Cao |
| 42 | Mọi sửa đổi ghi bằng Track Changes | Có, khi người dùng vào "tracked changes mode" [W] | **Có** từ bản 3.6 — mọi công cụ sửa nội dung và nút "Chèn vào chỗ con trỏ" luôn ghi có theo dõi, kể cả khi Track Changes đang tắt, xong trả lại chế độ cũ; Word dưới WordApi 1.4 thì từ chối ghi. (Bản 3.5 còn công tắc tắt nên chữ "mọi" khi đó là ghi quá.) | Chạy thật 15/11 (Track Changes tắt, nhờ AI chèn rồi thay chữ: Word hiện đúng phần thêm và phần xóa có theo dõi — ảnh 17). Tự kiểm tra bản 3.6.1: "sửa đổi vẫn được ghi (2 → 3), trả lại chế độ Off" (ảnh 21); bước này ở bản 3.6 đo sai | Cao |
| 43 | Xử lý nhận xét: sửa đúng chỗ được neo, trả lời nhận xét | Có [W] | **Có** | Tự kiểm tra | Cao |
| 44 | Tóm tắt sửa đổi của đối tác, phân mức độ | Có [W] | **Một phần** — đọc được và chấp nhận/từ chối từng sửa đổi; nhưng chỉ đọc tối đa 200 sửa đổi, không phân trang, không kèm số điều và câu trước/sau; phân mức độ hoàn toàn do model (Đợt 1A) | Tự kiểm tra (đọc 3 sửa đổi, chấp nhận 1) | Cao |
| 45 | Điền mẫu theo kiểu chữ của văn bản; điền bảng tại chỗ | Có [W] | **Một phần** — điền bảng tại chỗ: có (bảng đơn giản, không ô gộp). Theo kiểu chữ riêng của văn bản: chưa (chỉ dùng 8 kiểu dựng sẵn) | Tự kiểm tra (điền bảng) | Cao |
| 46 | Tìm mọi điều khoản theo chủ đề (không chỉ theo từ khóa) | Có [W] | **Một phần** — AI tự đọc rồi dẫn tới đoạn; văn bản dài phải đọc nhiều lượt, chưa có mục lục điều khoản | — | Cao |
| 47 | Đọc chú thích cuối trang, bookmark, đầu trang/chân trang | Có [W] | **Chưa** | — | Trung bình |
| 48 | Soát nhất quán: thuật ngữ, dẫn chiếu hỏng, đánh số thiếu, tên bên | Có (ví dụ sử dụng) [W] | **Một phần** — dựa vào model, chưa có công cụ dò riêng | — | Cao |
| 49 | Tệp .doc cũ | Không hỗ trợ [W] | Không hỗ trợ | — | — |
| | **POWERPOINT** | | | | |
| 50 | Tạo slide theo mẫu (slide master, bố cục, phông, màu) | Có [P] | **Một phần** — dùng bố cục và màu chủ đề; chưa đọc phông và slide master | Tự kiểm tra + Chạy thật | Cao (báo cáo tiến độ dự án) |
| 51 | Sửa đúng slide/hình được chọn | Có [P] | **Có** | Tự kiểm tra | Trung bình |
| 52 | Tạo cả bài từ mô tả | Có [P] | **Có** | Tự kiểm tra thao tác; chưa chạy thật trọn bài | Trung bình |
| 53 | Biểu đồ gốc PowerPoint, sửa được số liệu | Có [P] | **Có** — còn một lỗi: trên bài mới tinh, lần chèn biểu đồ đầu tiên làm mẫu thêm một bố cục (11 → 12); các lần sau dùng lại bố cục đó | Chạy thật; Tự kiểm tra 15/11: lần đầu trượt bước này, lần hai đạt | Cao |
| 54 | Chuyển gạch đầu dòng thành sơ đồ quy trình | Có [P] | **Có** | Tự kiểm tra | Trung bình |
| 55 | Sắp lại, gộp, xóa, nhân bản slide | Có [P] | **Có** (gộp: làm bằng nhiều bước) | Tự kiểm tra | Trung bình |
| | **KHẢ NĂNG MỚI CỦA BẢN CHÍNH THỨC (bổ sung 15/11/2026)** | | | | |
| 56 | Bộ nhớ dài hạn giữa các hội thoại | Có: Claude tự lưu ghi chú (vai trò, phản hồi, dự án), mỗi lần lưu hiện thông báo kèm Undo; chỉ khi đăng nhập tài khoản Claude và bật memory; không có trang xem/xóa [D] | **Chưa** (Đợt 5) | — | Trung bình |
| 57 | Tự kiểm hình thức sau khi sửa | Có ở Word và PowerPoint: add-in gọi thêm model để xem văn bản hoặc slide trông thế nào sau khi sửa [D] | **Chưa** (PowerPoint: Đợt 3, chụp ảnh slide; Word: Đợt 1A, đọc lại cấu trúc) | — | Cao |
| 58 | Soát lập luận trước khi sửa ngôn từ pháp lý | Có ở Word [D] | **Chưa** (Đợt 1B) | — | Cao |
| 59 | Tải bản ghi hội thoại | Có: ⋮ > Download Transcript [D][Pf] | **Chưa** (Đợt 5). Việc tải tệp từ khung add-in ra đĩa đã thử được 15/11 (tệp sao lưu) | — | Trung bình |
| 60 | Công cụ chạy mã Office.js do model viết | Có: nhật ký nêu công cụ `execute_office_js` bên cạnh công cụ chuyên biệt [OT] | **Chưa** — 54 công cụ khai báo sẵn; cân nhắc ở Đợt 5 (rủi ro cao, luôn hỏi kèm hiện mã) | — | Trung bình |
| 61 | Giới hạn thời gian và cỡ tệp | Có số cụ thể: chờ 90 giây cho bước chạy mã, 5 phút cho bước khác (web 2 phút), 30 giây cho một sửa chữ ở Word; mỗi lần đọc 2.000 ô có dữ liệu; rủi ro tăng khi Excel quá 1 triệu ô, PowerPoint từ 100 slide [Pf] | **Một phần** — giới hạn mỗi lần đọc (2.500 ô, văn bản dài đọc theo lượt); từ bản 3.7 có hạn chờ 120 giây cho từng bước công cụ; chưa cảnh báo tệp lớn | Tự kiểm tra (đọc văn bản dài 100 đoạn theo 2 lượt); hạn chờ: Kiểm thử máy | Trung bình |
| 62 | Phát hiện va chạm khi người dùng đang sửa | Có ở Excel, mức đo đếm: bộ đếm `office_agent.cell_edit_collision_total` ghi khi người dùng đang sửa ô lúc công cụ ghi [OT]; tài liệu không tả người dùng thấy gì. Kèm: phát hiện bước bị treo hoặc Office đổ giữa chừng (bản ghi cho mỗi bước đang chạy, quá 10 phút thì báo) [D][Pf] | **Chưa** | — | Thấp |
| 63 | Hội thoại gắn với tệp, tự mở khung cùng tài liệu | Có: ghi `claude.fileId` (nhận ra tệp sau khi đổi tên, Save As) và `Office.AutoShowTaskpaneWithDocument` vào tài liệu [D] | **Chưa** | — | Trung bình |
| 64 | Lấy lại phần ngữ cảnh đã đẩy ra | Có: phần bị đẩy khỏi ngữ cảnh làm việc được lưu riêng, Claude xin lại được [D] | **Chưa** — chỉ có tóm tắt hội thoại dài (dòng 9) | — | Cao (hợp đồng dài) |
| 65 | Dò khả năng của model khi kết nối | Có: gửi yêu cầu thử khi kết nối và khi đổi model để biết model, tìm web, chạy mã có dùng được [3P][D] | **Một phần** — nút Kiểm tra nhà cung cấp (bấm tay) và dấu 🛠 cho model gọi được công cụ; chưa tự dò khi đổi model | Chạy thật 15/11 (8 nguồn có key) | Cao |
| 66 | Bộ nhớ đệm lời nhắc ở mọi lượt | Có: add-in yêu cầu prompt caching ở mọi lần gọi [OT] | **Có** từ bản 3.7 với key Anthropic (mốc đệm ở lời nhắc hệ thống và ở cuối hội thoại; số token đọc từ bộ nhớ đệm hiện dưới mỗi câu trả lời). Nguồn khác: chỉ hiện số token đệm nếu nhà cung cấp tự báo | Chạy thật 15/11 (Word, Claude Haiku 5.5): lượt hai đọc 7.827 token từ bộ nhớ đệm; hội thoại dài đọc tới 483.000 token mỗi lượt từ bộ nhớ đệm | Cao (chi phí) |
| 67 | Phím tắt mở khung | Có: Ctrl+Alt+C, hành động ShowTaskpane [SJ] | **Chưa** — phải chuyển sang shared runtime; Đợt 5 | — | Thấp |

## Tổng hợp

- 67 dòng: **Có**: 27 · **Một phần**: 16 · **Chưa**: 20 · **Không làm được**: 1 · **Không cần / hai bên đều không**: 3. (Bản 3.6.1: 25 · 17 · 21 · 1 · 3; bản 3.7.0 chuyển dòng 5 từ Một phần lên Có và dòng 66 từ Chưa lên Có, cả hai đã chạy thật trong Office. Bản 14/11, 55 dòng: 29 · 11 · 11 · 1 · 3.)
- Ba khoảng trống ảnh hưởng nhất tới việc pháp lý – BĐS: số điều khoản tự đánh trong Word (#40), PDF quét (#10), ghi chéo ứng dụng (#20). Trong 12 dòng mới, ba dòng đáng làm sớm: soát lập luận trước khi sửa (#58), tự kiểm hình thức sau khi sửa (#57), lấy lại ngữ cảnh đã đẩy ra (#64).
- Điểm AI Đa Nguồn hơn: nhiều nhà cung cấp và model miễn phí; nút Hoàn tác riêng; nút Tự kiểm tra. Ngoài ra, ở chế độ third-party/gateway của bản chính thức (không dùng tài khoản Claude) chỉ **thiếu chéo ứng dụng và đọc giọng nói** (và không có thư mục connector); Skills dựng sẵn, Skills tải lên và tệp đính kèm vẫn có [3P]. AI Đa Nguồn có đọc chéo ứng dụng, chưa ghi chéo.

## Nguồn

Tài liệu chính thức, bản `.md`, tải từ máy ngày 15/11/2026:

- [X] https://claude.com/docs/office-agents/excel
- [W] https://claude.com/docs/office-agents/word
- [P] https://claude.com/docs/office-agents/powerpoint
- [O] https://claude.com/docs/office-agents/outlook
- [A] https://claude.com/docs/office-agents/work-across-apps
- [C] https://claude.com/docs/office-agents/connectors-and-skills
- [D] https://claude.com/docs/office-agents/data-storage
- [Dt] https://claude.com/docs/office-agents/dictation
- [3P] https://claude.com/docs/office-agents/third-party-platforms
- [Pf] https://claude.com/docs/office-agents/performance
- [OT] https://claude.com/docs/office-agents/opentelemetry
- Tổng quan: https://claude.com/docs/office-agents/overview · Quản trị: https://claude.com/docs/office-agents/enterprise-readiness
- [SJ] https://pivot.claude.ai/shortcuts.json (tải từ máy 15/11/2026: hành động `ShowTaskpane`, phím `Ctrl+Alt+C`)

Bốn bài hỗ trợ cũ (12650343, 13521390, 14465370, 13892150) đã gộp về các trang trên; ký hiệu [SX], [SP] của bản 14/11 không còn dùng.

Ghi chú: bản 14/11 ghi "đã bỏ phím tắt Ctrl+Alt+C vì không còn thấy trong tài liệu" là sai — phím tắt vẫn còn (dòng 67); trang tài liệu không nhắc, nhưng tệp shortcuts.json của add-in vẫn khai báo.
