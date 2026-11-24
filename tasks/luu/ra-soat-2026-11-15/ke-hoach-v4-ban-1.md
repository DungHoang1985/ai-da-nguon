# Kế hoạch hoàn thiện AI Đa Nguồn (từ v3.5 lên v4)

Lập ngày 14/11/2026. Phiên này chỉ lập kế hoạch, chưa sửa mã. Bảng so sánh với bản chính thức: `docs/so-sanh-claude-chinh-thuc.md`.

Trạng thái duyệt: xem mục "Quyết định" trong `tasks/HANDOFF.md`.

---

## A. Đánh giá kiến trúc

Đã kiểm trong phiên này (trên bản sao kho GitHub, nhánh `main`, commit 14/11/2026 00:03): `npm test` 71/71 đạt, kiểm kiểu TypeScript không lỗi, đóng gói thành công.

**Nhận định chung:** nền tốt, đáng giữ. Lõi tách khỏi Office nên kiểm thử được; ba kiểu API (OpenAI, Anthropic, Gemini) gom về một lớp; công cụ có cờ "ghi" và "rủi ro"; có Hoàn tác và nút Tự kiểm tra chạy thật trong Office. Không cần viết lại. Vấn đề nằm ở quy trình phát hành, vài lỗi thật và một số phần phình ra theo lịch sử.

Quy mô: 9.853 dòng (cả kiểm thử và cấu hình), 54 công cụ (Excel 20, Word 18, PowerPoint 14, dùng chung 2), 17 kỹ năng dựng sẵn.

### A1. Vấn đề cần xử lý

| # | Vấn đề | Mức | Bằng chứng | Cách xử lý | Đợt |
|---|---|---|---|---|---|
| 1 | **Thư mục `D:\ADD IN` không phải mã đang chạy.** Trong đó là bộ khung bản 0 (số nhận dạng add-in khác, README dạy gõ lệnh, chưa có chế độ tác tử). Mã v3.5 chỉ nằm trên GitHub. Ngược lại `docs/` và `tasks/` chỉ nằm trên máy. | Cao | `manifest.xml` trên máy: Id `ad31707e…`; trên kho: `c95713b4…`. Kho không có `docs/`, `tasks/`. | Lấy nhánh `main` làm nguồn chuẩn duy nhất; thay mã cũ trên máy; đưa `docs/`, `tasks/` vào kho. | 0 |
| 2 | **Phát hành bằng tay**: đóng gói rồi tự chép sang nhánh `gh-pages`. Dễ lệch, không có bước kiểm bắt buộc. | Cao | Không có thư mục `.github`; HANDOFF ghi bài học về `/tmp/pages`. | Tự động: đẩy lên `main` → chạy kiểm thử → đóng gói → phát hành. Kiểm thử trượt thì không phát hành. | 0 |
| 3 | **Dừng hoặc lỗi giữa chừng thì AI quên việc vừa làm.** Tài liệu đã bị sửa nhưng lịch sử gửi model không ghi lại lượt đó. | Cao | `Chat.tsx`: `historyRef.current = msgs` chỉ chạy khi `runAgent` kết thúc trọn vẹn. | Ghi lịch sử theo từng bước, kể cả khi dừng. | 0 |
| 4 | **Không tự thử lại** khi nhà cung cấp báo quá tải hoặc hết hạn mức tạm thời. Một lỗi 429 làm hỏng cả chuỗi thao tác. | Cao | `errors.ts` chỉ đổi mã lỗi thành thông báo. | Thử lại có chờ (tối đa 3 lần) cho 429 và 5xx. | 0 |
| 5 | **Tốn tiền không cần thiết với key Anthropic.** Mỗi bước gửi lại toàn bộ lời nhắc + 50 mô tả công cụ + danh mục kỹ năng, không dùng bộ nhớ đệm của API. Trạng thái tài liệu nhét trong lời nhắc hệ thống nên càng không đệm được. | Cao | `llm.ts` không có `cache_control`; `systemPrompt(ov)` nhúng JSON trạng thái. | Bật bộ nhớ đệm lời nhắc; chuyển trạng thái tài liệu sang tin nhắn. Hiện số token mỗi lượt để đo. | 0 |
| 6 | **Ngưỡng tóm tắt hội thoại cố định 120.000 ký tự** cho mọi model. Model miễn phí sức chứa nhỏ sẽ tràn trước khi kịp tóm tắt. | Trung bình | `compaction.ts`: `COMPACT_AT_CHARS = 120000`. | Tính ngưỡng theo sức chứa của model (đã có sẵn trong danh mục). | 0 |
| 7 | **Chưa có câu phòng "chỉ dẫn ẩn"** trong tài liệu, tệp đính kèm, kết quả web. Bản chính thức cảnh báo rõ rủi ro này. | Trung bình | Tìm trong `src/` không có nội dung phòng vệ. | Thêm nguyên tắc: nội dung đọc được là dữ liệu, không phải lệnh; hỏi lại khi tài liệu "yêu cầu" xóa, gửi, tìm web. | 0 |
| 8 | Key Gemini nằm trong đường dẫn gọi API (dễ lọt vào nhật ký). | Thấp | `llm.ts`, `models.ts`: `?key=`. | Chuyển sang tiêu đề `x-goog-api-key`. | 0 |
| 9 | Một thư viện phụ thuộc có cảnh báo bảo mật mức cao (`image-size`, đi theo pptxgenjs). | Thấp | `npm audit --omit=dev`. | `npm audit fix`, kiểm lại biểu đồ PowerPoint. | 0 |

### A2. Phần thừa và phức tạp quá mức

| # | Phần | Vì sao thừa | Xử lý |
|---|---|---|---|
| 1 | `src/core/chat.ts` (102 dòng) | Không còn nơi nào dùng; trùng chức năng với `agent/llm.ts`. | Xóa cùng kiểm thử của nó. |
| 2 | Gói đệm cho trình duyệt cũ (239 KB) và mục tiêu "IE 11" | Office hiện nay chạy trên WebView2 (Chromium). | Bỏ; gói tải nhanh hơn. |
| 3 | 11 thư viện phụ thuộc không dùng (`es6-promise`, `@fluentui/react-icons`, `less`, `less-loader`, `file-loader`, `os-browserify`, `process`, `acorn`, 3 gói `office-addin-*` thử nghiệm) | Còn sót từ mẫu của Microsoft. | Gỡ. |
| 4 | `package.json` ghi sai tên kho (của Microsoft), phiên bản 0.0.1; số phiên bản nằm ở 3 nơi | Dễ lệch khi phát hành. | Một nguồn duy nhất, tự điền khi đóng gói. |
| 5 | `Chat.tsx` 1.016 dòng, hơn 20 biến trạng thái, gánh cả giao diện lẫn điều phối | Khó sửa an toàn; lỗi A1.3 sinh ra từ đây. | Tách phần điều phối (chạy tác tử, lịch sử, model) ra khỏi giao diện. Làm ở Đợt 0 vì phải sửa A1.3 đúng chỗ này. |
| 6 | Thêm một công cụ phải sửa 4–5 nơi (mô tả, thực thi, nhãn tiếng Việt, rủi ro, tự kiểm tra); tệp chia "cơ bản/nâng cao" theo lịch sử bản 2, bản 3 | Dễ sót, dễ lệch. | Mỗi công cụ một khối khai báo đủ 5 phần. Chuyển dần: ứng dụng nào được đụng tới ở đợt nào thì chuyển ở đợt đó. |
| 7 | Hai cơ chế chọn kỹ năng (dò từ khóa + model tự gọi) | Chấp nhận được: dò từ khóa tiết kiệm một lượt gọi. | Giữ, không làm thêm. |

Không đề xuất: viết lại bằng khung khác, thêm máy chủ riêng, đổi thư viện giao diện.

---

## B. Đăng nhập thay cho dán API key

Nguồn đã đối chiếu ngày 14/11/2026, đường dẫn ở mục D.

| | (a) Đăng nhập OpenRouter một lần bấm | (b) Dùng gói Claude qua Agent SDK + chương trình trên máy | (c) Tín dụng API kèm gói Claude |
|---|---|---|---|
| **Có được phép không** | **Được.** OpenRouter công bố luồng này cho ứng dụng bên thứ ba; khóa nhận về thuộc tài khoản của người dùng. | **Vùng xám, không khuyến nghị.** Tài liệu của Anthropic: không cho nhà phát triển bên thứ ba đưa đăng nhập Claude.ai vào ứng dụng của họ; đăng nhập gói dành cho Claude Code và ứng dụng của Anthropic; ứng dụng dùng Agent SDK nên dùng API key. Tài liệu cũng nói không cấm một người tự đăng nhập Claude Code nguyên bản bằng gói của chính mình. Dùng riêng cho mình thì không bị cấm rõ ràng, nhưng Anthropic có quyền chặn mà không báo trước. Tuyệt đối không chia sẻ cho người khác. | **Được.** Đây là Claude API thông thường; Anthropic cấp tín dụng hằng tháng cho gói đủ điều kiện. |
| **Thực chất là gì** | Bấm nút → mở trang OpenRouter → bấm Authorize → add-in tự nhận khóa. Không phải dán. | Add-in gửi yêu cầu tới một chương trình chạy nền trên máy; chương trình gọi Claude Code đã đăng nhập gói. | Vẫn là API key Anthropic, dán một lần. Không phải đăng nhập, nhưng không tốn thêm tiền trong mức tín dụng. |
| **Chi phí** | Trả theo lượng dùng vào số dư OpenRouter; có model miễn phí. | Không thêm tiền; trừ vào hạn mức gói (chung với Claude anh đang dùng hằng ngày). | 0 đồng trong mức: Max 100 USD/tháng, Ultra 200 USD/tháng, Team 20 hoặc 100 USD mỗi chỗ (tối đa 500). **Pro, Free, Enterprise không có.** Không cộng dồn sang tháng sau. Hết thì dừng, không trừ vào gói. |
| **Model** | Hầu hết các hãng, gồm Claude. | Chỉ Claude. | Chỉ Claude. |
| **Rủi ro** | Khóa lưu trên máy như hiện nay. Phụ thuộc OpenRouter. Cần thử kỹ thuật: luồng chạy được trong hộp thoại của Office. | Chính sách đã đổi qua lại trong năm 2026; có thể bị chặn. Phải cài và giữ một chương trình luôn chạy. Phải dựng lại cách add-in điều khiển công cụ. Tốn công nhất trong ba hướng. | Gần như không. Phải ở gói đủ điều kiện ít nhất 7 ngày; mỗi gói chỉ liên kết một tổ chức Console và không tự đổi được. |
| **Công sức làm** | Nhỏ (một màn hình, một trang nhận kết quả). | Lớn. | Không cần sửa mã; thêm một trang hướng dẫn và cảnh báo hết tín dụng. |
| **Anh phải làm gì** | Có tài khoản OpenRouter; bấm Đăng nhập, bấm Authorize. Muốn dùng model trả phí thì nạp tiền ở OpenRouter. | Để Claude cài chương trình; đăng nhập Claude Code một lần. | Vào claude.ai → Cài đặt → Thanh toán → mục tín dụng API → liên kết tổ chức → xác nhận điều khoản. Claude thao tác cùng anh; anh bấm xác nhận. |

Lưu ý về bài hỗ trợ 15036540: bài này nói về dùng Agent SDK với gói Claude. Nội dung hiện tại: từ 07/10/2026 gói Max, Ultra, Team có tín dụng API hằng tháng (chi tiết ở bài 17154008); khi đăng nhập bằng gói thì Agent SDK và `claude -p` vẫn trừ vào hạn mức gói. Khoản "tín dụng Agent SDK" công bố tháng 6/2026 đã bị hủy.

**Khuyến nghị:** làm (a), và dùng (c) nếu anh ở gói Max trở lên. Không làm (b).

Phát hiện thêm: add-in Claude chính thức có sẵn lối "Cloud provider or gateway" cho phép nhập địa chỉ cổng trung gian và token thay cho tài khoản Claude. Lối này cần một cổng trung gian chạy riêng và mất Skills, tệp đính kèm, chéo ứng dụng, đọc giọng nói. Không thay được AI Đa Nguồn; chỉ đáng dùng làm mốc so chất lượng thao tác.

---

## C. Kế hoạch theo đợt

Nguyên tắc: mỗi đợt xong là dùng được ngay; Claude tự làm và tự kiểm trên máy anh (điều khiển Office như các bản trước); anh không gõ lệnh. Bước nào không đạt thì làm lại riêng bước đó, tối đa 3 lần, rồi dừng báo cáo. Cuối mỗi đợt có bảng Tiêu chí · Kết quả · Bằng chứng.

Các mục ghi "(thử trước)" là chưa chắc về kỹ thuật: Claude thử trong 1 buổi; không được thì báo ngay và chuyển phương án dự phòng, không cố.

### Đợt 0 — Dọn nền và kiểm nốt (bản 3.6)

Không thêm tính năng lớn. Làm trước vì mọi đợt sau dựa vào.

Việc:
1. Một nguồn mã duy nhất + phát hành tự động (A1.1, A1.2).
2. Sửa lỗi: quên việc khi dừng giữa chừng, tự thử lại, ngưỡng tóm tắt theo model (A1.3, A1.4, A1.6); tách phần điều phối khỏi `Chat.tsx` (A2.5).
3. Giảm chi phí với key Anthropic, hiện số token mỗi lượt (A1.5).
4. Phòng "chỉ dẫn ẩn", key Gemini, cập nhật thư viện (A1.7–A1.9).
5. Gỡ phần thừa (A2.1–A2.4).
6. Thêm nút **Kiểm tra nhà cung cấp**: với mỗi key đã nhập, chạy thật một chuỗi "model gọi công cụ → nhận kết quả → trả lời".
7. Kiểm nốt 7 mục "Chưa kiểm được" của bản 3.

| Tiêu chí nghiệm thu | Cách tự kiểm trên máy anh |
|---|---|
| Đầu khung hiện v3.6; `D:\ADD IN` trùng với kho | So mã băm của `main` với thư mục; chụp đầu khung |
| Đẩy mã lên là tự phát hành; kiểm thử trượt thì không phát hành | Ảnh chụp lần chạy tự động thành công và một lần cố ý trượt |
| Tự kiểm tra vẫn đạt Excel 22/22, Word 14/14, PowerPoint 14/14 | Bấm nút ở từng ứng dụng, chụp kết quả |
| Bấm Dừng giữa chuỗi sửa, hỏi "anh vừa làm gì" → AI kể đúng phần đã làm | Chạy thật trong Word |
| Lượt thứ hai của cùng hội thoại (key Anthropic) có token đọc từ bộ nhớ đệm lớn hơn 0 | Số token hiện ngay trong khung |
| 7 mục chưa kiểm: tìm web; mở lại và xóa hội thoại cũ; tự tóm tắt hội thoại dài; bấm địa chỉ ô/đoạn/slide; nhập kỹ năng qua giao diện; đính kèm ảnh; lưu Chỉ dẫn riêng | Mỗi mục một ảnh chụp kết quả thật; mục nào trượt thì sửa trong đợt |
| Mỗi nhà cung cấp có key: nút Kiểm tra nhà cung cấp báo Đạt hoặc nêu rõ lý do | Chạy với key anh có; nguồn chưa có key ghi "Chưa kiểm được" |

### Đợt 1 — Word cho việc pháp lý (bản 4.0)

Giá trị cao nhất với anh. Lấp các dòng 10, 40, 46, 47, 48 của bảng so sánh.

Việc:
1. Đọc **số điều khoản do Word tự đánh**, cấp của điều khoản; lập mục lục điều khoản để AI tra nhanh trên văn bản dài.
2. Đọc chú thích cuối trang, đầu trang/chân trang, bookmark.
3. Công cụ soát nhất quán: thuật ngữ định nghĩa dùng lệch, dẫn chiếu tới điều không tồn tại, đánh số nhảy cóc, tên bên gọi khác nhau.
4. **PDF quét**: gửi nguyên tệp cho model đọc được ảnh (Anthropic, Gemini); nguồn khác thì báo rõ không đọc được.
5. Công tắc "Suy nghĩ kỹ" cho việc rà soát (model nào hỗ trợ thì bật).
6. Chuyển công cụ Word sang cách khai báo một khối (A2.6).

| Tiêu chí nghiệm thu | Cách tự kiểm trên máy anh |
|---|---|
| Hợp đồng thử từ 20 trang, đánh số tự động 3 cấp: 10 câu hỏi "Điều x.y nói gì / quy định về … nằm ở điều nào" đúng số điều từ 9/10 | Claude tạo hợp đồng thử, chạy 10 câu, lập bảng đối chiếu |
| Cài sẵn 3 thuật ngữ dùng lệch, 2 dẫn chiếu hỏng, 1 chỗ đánh số nhảy: công cụ soát tìm đủ 6 | Chạy trên hợp đồng thử, chụp kết quả |
| PDF quét 3 trang: AI tóm tắt đúng nội dung và trích đúng 3 số liệu | Chạy với key Anthropic; Gemini nếu có key |
| Chú thích cuối trang và chân trang có trong câu trả lời khi được hỏi | Chạy thật |
| Tự kiểm tra Word thêm các bước mới, đạt hết; các bước cũ không hỏng | Bấm nút, chụp kết quả |
| Sửa đổi vẫn ghi bằng Track Changes, số điều khoản không vỡ | Kiểm trên hợp đồng thử sau 5 lần sửa |

### Đợt 2 — Đăng nhập và chi phí (bản 4.1)

Việc:
1. Nút **Đăng nhập OpenRouter** (thử trước: luồng trong hộp thoại Office). Dự phòng: OpenRouter hiện mã, anh dán mã một lần.
2. Nếu anh ở gói Max trở lên: nhận tín dụng API theo gói, tạo key, nhập vào add-in.
3. Cảnh báo khi số dư thấp hoặc hết tín dụng, gợi ý đổi sang nguồn khác.

| Tiêu chí nghiệm thu | Cách tự kiểm trên máy anh |
|---|---|
| Từ trạng thái chưa có key OpenRouter: bấm một nút → đăng nhập → danh sách model hiện → sửa được một đoạn Word | Xóa key OpenRouter, làm lại từ đầu, quay lại các bước |
| Đăng xuất xóa sạch khóa trên máy | Kiểm trong Cài đặt và bộ nhớ của add-in |
| (Nếu Max) trang Thanh toán của Console hiện tín dụng; một yêu cầu chạy xong, số dư giảm | Ảnh chụp trước và sau |
| Khi nhà cung cấp báo hết tiền: thông báo tiếng Việt rõ, không mất hội thoại | Thử với một key không có số dư |

### Đợt 3 — Chéo ứng dụng thật và báo cáo tiến độ (bản 4.2)

Lấp các dòng 20, 21, 50.

Việc:
1. (Thử trước) Khung ở ứng dụng này **gửi lệnh** cho khung ở ứng dụng kia thực hiện: AI ở Word ghi được sang Excel/PowerPoint đang mở. Luôn hỏi xác nhận trước khi ghi sang ứng dụng khác.
2. Chỉ báo ứng dụng nào đang kết nối; mang tóm tắt hội thoại sang khi đổi ứng dụng.
3. PowerPoint: đọc phông và slide master; cập nhật biểu đồ đã có theo số mới.

| Tiêu chí nghiệm thu | Cách tự kiểm trên máy anh |
|---|---|
| Từ PowerPoint: "lấy bảng tiến độ ở Excel đang mở, tạo 3 slide và 1 biểu đồ" chạy trọn, đúng số | Chạy thật với tệp tiến độ thử |
| Từ Excel: "ghi tóm tắt sang Word đang mở" → đoạn mới xuất hiện trong Word, có hỏi xác nhận, hoàn tác được ở Word | Chạy thật |
| Đóng khung ở ứng dụng kia → AI báo rõ không ghi được, không treo | Chạy thật |
| Slide mới dùng đúng phông của mẫu | So với slide master |

Dự phòng nếu mục 1 không làm được: giữ đọc chéo, thêm nút "Chuyển sang … và làm tiếp" kèm sẵn nội dung.

### Đợt 4 — Excel và dấu vết soát lại (bản 4.3)

Lấp các dòng 30–33, 35, 36.

Việc: sheet "Nhật ký AI" ghi từng thay đổi; tô sáng ô vừa sửa (tắt được); sửa biểu đồ và PivotTable có sẵn; vùng in; hỏi trước khi dùng hàm nguy hiểm.

| Tiêu chí nghiệm thu | Cách tự kiểm trên máy anh |
|---|---|
| Sau 5 thao tác ghi: sheet nhật ký có đủ 5 dòng (giờ, ô, giá trị cũ, giá trị mới, lý do) | Tự kiểm tra + chạy thật |
| Hoàn tác thì ô hết tô sáng và nhật ký ghi thêm dòng hoàn tác | Chạy thật |
| Đổi kiểu và tiêu đề một biểu đồ có sẵn; thêm một trường vào Pivot có sẵn | Tự kiểm tra |
| Công thức có WEBSERVICE hoặc INDIRECT → có hỏi trước khi ghi | Tự kiểm tra |

### Đợt 5 — Tùy chọn, làm khi anh cần

Đọc một trang web cụ thể; connectors qua tính năng MCP của Claude API (thử trước, chỉ với key Anthropic); ba mức quyền như bản chính thức; chạy mã; Outlook.

### Đưa vào kho `D:\CLAUDE`

Khi đóng dự án: lập phiếu PĐX để đưa vào kho (a) bảng so sánh và cách kiểm "Tự kiểm tra / Chạy thật / Kiểm thử máy", (b) các bài học Office.js trong HANDOFF, (c) bộ hợp đồng thử và 10 câu hỏi kiểm của Đợt 1 (dùng lại được cho hệ thống tra cứu pháp lý).

---

## D. Nguồn đã đối chiếu

Add-in chính thức: xem cuối `docs/so-sanh-claude-chinh-thuc.md`.

Đăng nhập:
- OpenRouter OAuth PKCE: https://openrouter.ai/docs/guides/overview/auth/oauth
- Dùng Agent SDK với gói Claude (bài 15036540): https://support.claude.com/en/articles/15036540
- Tín dụng API hằng tháng cho Max, Ultra, Team (bài 17154008): https://support.claude.com/en/articles/17154008
- Quy định về đăng nhập và thông tin xác thực: https://code.claude.com/docs/en/legal-and-compliance
- Tổng quan Agent SDK (ghi chú về đăng nhập claude.ai): https://code.claude.com/docs/en/agent-sdk/overview
- Kết nối add-in chính thức qua cổng trung gian: https://claude.com/docs/office-agents/third-party-platforms
