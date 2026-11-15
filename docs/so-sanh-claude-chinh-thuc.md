# So sánh AI Đa Nguồn (v3.5) với add-in Claude chính thức

Cập nhật 14/11/2026. Thay toàn bộ bản cũ (các mục 1–7 trước đây đã gộp vào một bảng dưới đây).

## Cách đọc cột "Bằng chứng"

- **Tự kiểm tra**: nút Tự kiểm tra chạy thao tác thật trong Office trên máy anh (bản 3.4–3.5: Excel 22/22, Word 14/14, PowerPoint 14/14 — theo HANDOFF).
- **Chạy thật**: đã giao việc cho AI thật (Claude Opus 5.5, key Anthropic) và thấy kết quả trong tài liệu.
- **Kiểm thử máy**: chỉ có kiểm thử tự động ngoài Office (`npm test` 71/71, chạy lại 14/11/2026). Chưa thấy chạy trong Office.
- **Giao diện**: thấy hiện đúng trên màn hình, chưa thao tác thật.
- **—**: chưa có gì để kiểm.

Lưu ý chung: mọi mục "Chạy thật" mới chỉ chạy với key Anthropic. OpenRouter, Gemini, Groq và các nguồn khác **chưa chạy thật** lần nào.

Nguồn (viết tắt ở cột 2): [X] Excel · [W] Word · [P] PowerPoint · [A] Work across apps · [C] Connectors and Skills · [D] Data storage · [Dt] Dictation · [3P] Third-party platforms · [SX] [SP] bài hỗ trợ Excel, PowerPoint. Đường dẫn ở cuối tệp.

## Bảng so sánh

| # | Khả năng | Bản chính thức | AI Đa Nguồn | Bằng chứng | Mức quan trọng (pháp lý – BĐS) |
|---|---|---|---|---|---|
| | **CHUNG** | | | | |
| 1 | Đăng nhập bằng tài khoản Claude, dùng hạn mức của gói | Có [X][W][P] | **Không làm được** (chỉ ứng dụng của Anthropic được dùng đăng nhập này; xem kế hoạch, phần đăng nhập) | — | Cao |
| 2 | Kết nối bằng token/key, không cần tài khoản Claude | Có, qua "Cloud provider or gateway" [3P] | **Có** — nhập key của 11 nhà cung cấp | Chạy thật (Anthropic); nguồn khác chưa | Cao |
| 3 | Chọn model | Có, danh sách Claude chọn lọc [X] | **Có**, rộng hơn (cả model miễn phí) | Chạy thật (Anthropic) | Cao |
| 4 | AI tự đọc và sửa tài liệu qua nhiều bước | Có | **Có** (54 công cụ) | Tự kiểm tra + Chạy thật | Cao |
| 5 | Hỏi xác nhận trước thao tác rủi ro; cảnh báo trước khi ghi đè | Có [X] | **Có** | Tự kiểm tra | Cao |
| 6 | Ba mức quyền: hỏi trước khi sửa / chấp nhận mọi sửa / bỏ qua mọi xin phép | Có [3P] | **Một phần** — có công tắc "hỏi trước mọi thao tác" và "cho phép hết lần này" | Tự kiểm tra | Trung bình |
| 7 | Chỉ dẫn riêng lưu sẵn cho từng ứng dụng | Có [X][W][P] | **Có** | Giao diện (chưa lưu thử) | Cao |
| 8 | Lịch sử hội thoại lưu trên máy, xóa được | Có, 50 hội thoại gần nhất [D] | **Có**, 60 hội thoại | Còn nguyên sau khi tải lại; mở lại/xóa: Kiểm thử máy | Trung bình |
| 9 | Tự thu gọn hội thoại dài | Có [X] | **Có** — nhưng ngưỡng cố định, không theo sức chứa của từng model | Kiểm thử máy | Trung bình |
| 10 | Tệp đính kèm làm nguồn | Có (không có ở chế độ gateway) [3P] | **Một phần** — PDF có lớp chữ, DOCX, TXT/CSV/MD/JSON, ảnh. **PDF quét (scan) chưa đọc được** | Chạy thật (PDF + DOCX); ảnh: Kiểm thử máy | **Cao** — hồ sơ pháp lý phần lớn là bản quét |
| 11 | Kỹ năng (Skills) tự áp dụng theo việc | Có [C] | **Có** — 17 kỹ năng dựng sẵn | Chạy thật (1 kỹ năng); còn lại Kiểm thử máy | Cao |
| 12 | Gọi kỹ năng bằng "/", lọc theo ứng dụng | Có [C] | **Có** | Giao diện (Excel) | Trung bình |
| 13 | Tải kỹ năng riêng lên, lưu trên máy | Có [D] | **Có** (SKILL.md, .zip, .skill, gói plugin) | Kiểm thử máy | Cao |
| 14 | Connectors (S&P, LSEG, Daloopa, MCP tùy biến) | Có, nút "+" [C] | **Chưa** — với key Anthropic có thể làm một phần, cần thử kỹ thuật | — | Thấp |
| 15 | Tìm kiếm web | Có [3P] | **Một phần** — theo nhà cung cấp (OpenRouter, Anthropic, Gemini) | Kiểm thử máy | Cao (tra cứu văn bản) |
| 16 | Đọc nội dung một trang web cụ thể | Có [3P] | **Chưa** | — | Trung bình |
| 17 | Chạy mã để tính toán, phân tích | Có [3P] | **Chưa** | — | Thấp |
| 18 | Đọc giọng nói | Có, chỉ bản máy tính [Dt] | **Chưa** — dùng Win+H của Windows thay được | — | Thấp |
| 19 | Chéo ứng dụng: **đọc** tệp đang mở ở ứng dụng khác | Có [A] | **Có** — qua ảnh chụp nội dung, tối đa 30 phút tuổi | Chạy thật (PowerPoint đọc Excel) | Cao |
| 20 | Chéo ứng dụng: **ghi** sang tệp đang mở ở ứng dụng khác | Có [A] | **Chưa** | — | Cao (Excel → Word/PowerPoint báo cáo) |
| 21 | Hội thoại đi theo khi đổi ứng dụng; chỉ báo ứng dụng đang kết nối | Có [A] | **Chưa** | — | Trung bình |
| 22 | Trích dẫn bấm được (ô, đoạn, slide) | Có [X][W] | **Có** | Kiểm thử máy (chưa bấm trong Office) | Cao |
| 23 | Outlook | Có (bản thử) | **Chưa**, ngoài phạm vi | — | Thấp |
| 24 | Quản trị doanh nghiệp, Compliance API, nhật ký OpenTelemetry | Có | **Không cần** (dùng cá nhân) | — | Thấp |
| | **EXCEL** | | | | |
| 25 | Hỏi đáp nhiều sheet, giải thích công thức | Có [X] | **Có** | Tự kiểm tra | Trung bình |
| 26 | Đổi giả định, giữ nguyên quan hệ công thức | Có [X] | **Có** | Tự kiểm tra (ghi ô); chưa chạy thật trên mô hình | Trung bình |
| 27 | Tìm gốc lỗi #REF!, #DIV/0!, lần nguồn số liệu | Có [X] | **Có** (quét lỗi, lần ô trước/sau); tham chiếu vòng: chưa | Tự kiểm tra | Trung bình |
| 28 | Tạo mô hình, điền mẫu có sẵn | Có [X] | **Có** | Tự kiểm tra thao tác; chưa chạy thật | Trung bình |
| 29 | Sắp xếp, lọc, định dạng có điều kiện, danh sách chọn | Có [X] | **Có** | Tự kiểm tra | Trung bình |
| 30 | PivotTable | Có, tạo và sửa [X] | **Một phần** — tạo mới; chưa sửa Pivot có sẵn | Tự kiểm tra | Thấp |
| 31 | Biểu đồ | Có, tạo và sửa [SX] | **Một phần** — tạo mới; chưa sửa biểu đồ có sẵn | Tự kiểm tra | Trung bình |
| 32 | Tô sáng ô vừa sửa, kèm ghi chú giải thích | Có [SX] | **Một phần** — bấm vào bước để nhảy tới ô; không tô màu | Tự kiểm tra | Trung bình |
| 33 | Nhật ký phiên ghi vào một sheet ("Claude Log") | Có [SX] | **Chưa** | — | Trung bình (dấu vết để soát lại) |
| 34 | Lấy bảng từ PDF, kể cả bản quét, đưa vào mẫu | Có [SX] | **Một phần** — chỉ PDF có lớp chữ | Chạy thật (đọc PDF chữ) | Cao |
| 35 | Chuẩn bị in (đường lưới, vùng in) | Có [SX] | **Chưa** | — | Thấp |
| 36 | Hỏi trước khi dùng hàm nguy hiểm (WEBSERVICE, INDIRECT…) | Có [SX] | **Chưa** | — | Thấp |
| 37 | Hoàn tác | Ctrl+Z của Excel [SX] | **Có, hơn** — nút Hoàn tác riêng | Tự kiểm tra + Chạy thật | Cao |
| 38 | Macro/VBA, Data table | Không hỗ trợ [X] | Không hỗ trợ | — | — |
| | **WORD** | | | | |
| 39 | Hỏi đáp về văn bản, trích dẫn mục bấm được | Có [W] | **Có** (dạng [¶ số đoạn]) | Tự kiểm tra (đọc); bấm: Kiểm thử máy | Cao |
| 40 | Hiểu đánh số pháp lý nhiều cấp, thuật ngữ định nghĩa, dẫn chiếu chéo | Có [W] | **Một phần** — **chưa đọc số điều khoản do Word tự đánh**, nên AI có thể không biết đoạn nào là "Điều 5.3" | — | **Cao nhất** |
| 41 | Sửa đoạn đang chọn, giữ kiểu chữ và đánh số | Có [W] | **Có** | Tự kiểm tra | Cao |
| 42 | Mọi sửa đổi ghi bằng Track Changes | Có [W] | **Có** | Tự kiểm tra | Cao |
| 43 | Xử lý nhận xét: sửa đúng chỗ được neo, trả lời nhận xét | Có [W] | **Có** | Tự kiểm tra | Cao |
| 44 | Tóm tắt sửa đổi của đối tác, phân mức độ | Có [W] | **Có**, thêm chấp nhận/từ chối từng sửa đổi | Tự kiểm tra | Cao |
| 45 | Điền mẫu theo kiểu chữ của văn bản; điền bảng tại chỗ | Có [W] | **Có** | Tự kiểm tra | Cao |
| 46 | Tìm mọi điều khoản theo chủ đề (không chỉ theo từ khóa) | Có [W] | **Một phần** — AI tự đọc rồi dẫn tới đoạn; văn bản dài phải đọc nhiều lượt, chưa có mục lục điều khoản | — | Cao |
| 47 | Đọc chú thích cuối trang, bookmark, đầu trang/chân trang | Có [W] | **Chưa** | — | Trung bình |
| 48 | Soát nhất quán: thuật ngữ, dẫn chiếu hỏng, đánh số thiếu, tên bên | Có (ví dụ sử dụng) [W] | **Một phần** — dựa vào model, chưa có công cụ dò riêng | — | Cao |
| 49 | Tệp .doc cũ | Không hỗ trợ [W] | Không hỗ trợ | — | — |
| | **POWERPOINT** | | | | |
| 50 | Tạo slide theo mẫu (slide master, bố cục, phông, màu) | Có [P] | **Một phần** — dùng bố cục và màu chủ đề; chưa đọc phông và slide master | Tự kiểm tra + Chạy thật | Cao (báo cáo tiến độ dự án) |
| 51 | Sửa đúng slide/hình được chọn | Có [P] | **Có** | Tự kiểm tra | Trung bình |
| 52 | Tạo cả bài từ mô tả | Có [P] | **Có** | Tự kiểm tra thao tác; chưa chạy thật trọn bài | Trung bình |
| 53 | Biểu đồ gốc PowerPoint, sửa được số liệu | Có [P] | **Có** | Chạy thật | Cao |
| 54 | Chuyển gạch đầu dòng thành sơ đồ quy trình | Có [P] | **Có** | Tự kiểm tra | Trung bình |
| 55 | Sắp lại, gộp, xóa, nhân bản slide | Có [P] | **Có** (gộp: làm bằng nhiều bước) | Tự kiểm tra | Trung bình |

## Tổng hợp

- **Có**: 29 · **Một phần**: 11 · **Chưa**: 11 · **Không làm được**: 1 · **Không cần / hai bên đều không**: 3.
- Ba khoảng trống ảnh hưởng nhất tới việc pháp lý – BĐS: số điều khoản tự đánh trong Word (#40), PDF quét (#10, #34), ghi chéo ứng dụng (#20).
- Điểm AI Đa Nguồn hơn: nhiều nhà cung cấp và model miễn phí; nút Hoàn tác riêng; nút Tự kiểm tra. Ngoài ra, ở chế độ gateway của bản chính thức (không dùng tài khoản Claude) thì Skills, tệp đính kèm, chéo ứng dụng và đọc giọng nói đều **không có** [3P] — AI Đa Nguồn có ba mục đầu.

## Nguồn

- [X] https://claude.com/docs/office-agents/excel
- [W] https://claude.com/docs/office-agents/word
- [P] https://claude.com/docs/office-agents/powerpoint
- [A] https://claude.com/docs/office-agents/work-across-apps
- [C] https://claude.com/docs/office-agents/connectors-and-skills
- [D] https://claude.com/docs/office-agents/data-storage
- [Dt] https://claude.com/docs/office-agents/dictation
- [3P] https://claude.com/docs/office-agents/third-party-platforms
- Tổng quan: https://claude.com/docs/office-agents/overview
- [SX] https://support.claude.com/en/articles/12650343
- [SP] https://support.claude.com/en/articles/13521390
- Word (bài hỗ trợ): https://support.claude.com/en/articles/14465370
- Chéo ứng dụng (bài hỗ trợ): https://support.claude.com/en/articles/13892150

Ghi chú: phím tắt Ctrl+Alt+C ở bản cũ đã bỏ khỏi bảng vì không còn thấy trong tài liệu chính thức hiện tại.
