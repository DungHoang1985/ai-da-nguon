# Kế hoạch hoàn thiện AI Đa Nguồn — bản 2

Lập ngày 15/11/2026. **Trạng thái: ĐÃ DUYỆT ngày 15/11/2026.** Thay bản sáng 14/11 (lưu tại `tasks/ra-soat-2026-11-15/ke-hoach-v4-ban-1.md`). Việc cụ thể của đợt đang làm: `tasks/dot-0.md` (Đợt 0A).

Quyết định của anh ngày 15/11/2026:
1. Duyệt kế hoạch bản 2, giữ thứ tự 0A → 0B → 1A → 1B → 2 → 3 → 4.
2. Bố cục khung chat: kiểu **một dòng đầu** (bỏ hai thẻ Chat / Cài đặt; nhãn model cạnh ô nhập; Cài đặt mở bằng nút bánh răng).
3. Lỗi đổi nhà cung cấp: anh không nhớ rõ cặp nguồn nào → Claude tự thử mọi cặp nguồn anh có key trong Office.
4. **Bỏ công tắc tắt Track Changes**: sửa đổi của AI trong Word luôn được ghi có theo dõi.

---

## 1. Tóm tắt

Kế hoạch sáng 14/11 đúng hướng. Giữ nguyên các quyết định anh đã chọn: không viết lại, không thêm máy chủ; làm Word pháp lý trước đăng nhập; đăng nhập = OpenRouter một lần bấm + tín dụng API kèm gói Max; không dùng gói Claude qua Agent SDK.

Bản 2 thay đổi 7 điểm:

1. **Đợt 0 tách đôi.** 0A gồm nền phát hành và bốn việc anh nêu (Enter để gửi, bố cục gọn, lỗi đổi nhà cung cấp, sao lưu) để anh thấy khác ngay. 0B gồm độ bền, chi phí, an toàn.
2. **Thêm việc mức cao: sao lưu và khôi phục dữ liệu add-in.** Office lưu dữ liệu add-in theo từng tài khoản Office; máy anh đang có hai kho. Đổi tài khoản hoặc sau một bản cập nhật Office, key và lịch sử có thể "biến mất".
3. **Lỗi mất kết nối khi đổi nhà cung cấp đã khoanh được nguyên nhân** (mục 3.3), chưa tái hiện trong Office.
4. **Đợt 1 (Word) rộng hơn.** Mã Word thiếu năm việc nền mà kế hoạch cũ chưa có: chống sửa nhầm đoạn, sửa ở mức từ, đọc đúng khi văn bản đang có sửa đổi của đối tác, đọc văn bản dài, Hoàn tác cho Word. Việc đọc số điều khoản phải thăm dò trong Word thật trước khi cam kết.
5. **Bài nghiệm thu Word lấy từ ví dụ sử dụng chính thức**, chạy song song trên add-in Claude chính thức (anh có gói Max) để so trực tiếp.
6. **Bảng so sánh 55 dòng có 3 chỗ sai và thiếu khoảng 12 khả năng mới** của bản chính thức (mục 7).
7. **Chéo ứng dụng:** bản chính thức làm qua máy chủ của Anthropic. Mình không có máy chủ nên phải thử kênh cục bộ; thêm bước kiểm ba ứng dụng có dùng chung một kho lưu trữ không.

---

## 2. Hiện trạng đã kiểm lại

Ba mức bằng chứng dùng trong tệp này: **Tôi kiểm** (tôi chạy lệnh hoặc đọc mã trực tiếp) · **Tác tử** (tác tử rà soát báo cáo, có dòng mã hoặc đường dẫn nguồn) · **Chưa thử trong Office**.

| Điều kiểm | Kết quả | Bằng chứng |
|---|---|---|
| Mã `main` trên GitHub | Vẫn là commit `a38ee20` (14/11 00:03), không đổi từ lần audit trước | Tôi kiểm: `git log` trên bản sao |
| Kiểm thử, kiểm kiểu, đóng gói **trên máy anh** | `npm test` 71/71, `tsc` không lỗi, đóng gói đạt, 1 cảnh báo bảo mật mức cao (`image-size`) | Tôi kiểm: chạy trên bản sao ở vùng tạm |
| Máy anh có Git, Node, quyền đẩy lên kho | Git 2.56, Node 24.14, npm 11.9; `git push --dry-run` thành công, không cần đăng nhập lại | Tôi kiểm |
| Office | Microsoft 365 bản 16.0.20527 (Current Channel), WebView2 157: vượt mọi ngưỡng API cần dùng | Tôi kiểm: registry (chỉ đọc) |
| Nơi lưu dữ liệu add-in | Hai kho tách theo hai tài khoản Office (một tài khoản tổ chức, một tài khoản cá nhân) | Tôi kiểm: liệt kê thư mục `Wef\webview2` |
| Chín vấn đề A1 của kế hoạch cũ | A1.3, A1.4, A1.5, A1.6, A1.8 đúng; A1.9 đúng; A2.2, A2.4 đúng | Tôi kiểm: đọc `loop.ts`, `llm.ts`, `compaction.ts`, `errors.ts`, `Chat.tsx`, `webpack.config.js` |
| Gửi tin nhắn | Chỉ Ctrl+Enter mới gửi | Tôi kiểm: `Chat.tsx:974–976` |
| Vùng chat | Phía trên luôn có 5 hàng điều khiển + hàng tiêu đề + hàng thẻ | Tôi kiểm: `Chat.tsx:716–782`, `App.tsx:35–46` |
| Nhà cung cấp chặn gọi từ add-in | SambaNova và NVIDIA không trả tiêu đề cho phép gọi từ trình duyệt; 9 nguồn còn lại cho phép | Tác tử đo ngày 14/11 (tệp kết quả thô còn giữ) |
| Mã Word | 18 công cụ, 20 phát hiện (12 mức cao); 5 dòng của bảng so sánh ghi quá mức thực tế | Tác tử, có dòng mã; phần "đã tái hiện" là chạy logic thật ngoài Office |

**Phạm vi chưa rà trong phiên này:** mã Excel, mã PowerPoint, phần kỹ năng và tệp đính kèm chưa được rà sâu (10/14 tác tử bị dừng vì hết hạn mức model). Sẽ rà khi viết `tasks/dot-N.md` của đợt tương ứng, trước khi đụng mã đợt đó.

---

## 3. Bốn việc anh nêu

### 3.1 Enter để gửi
Enter gửi; Shift+Enter xuống dòng; Ctrl+Enter vẫn gửi. Không gửi khi bộ gõ tiếng Việt đang ghép chữ. Khi danh sách "/" đang mở: Enter chọn mục, mũi tên di chuyển, Esc đóng. Đang chạy thì Esc = Dừng.

### 3.2 Bố cục: phần cài đặt ẩn vào Cài đặt, vùng chat rộng
Theo tài liệu chính thức, khung Claude chỉ có vùng hội thoại và ô nhập với một hàng nút bên dưới (nút "+", mức quyền, micro, chọn model); mọi thứ cấu hình một lần nằm trong Settings.

Đề xuất cho AI Đa Nguồn:
- **Đầu khung một dòng:** tên + số phiên bản nhỏ · Hội thoại mới · Lịch sử · Cài đặt (⚙). Bỏ thanh thẻ Chat / Cài đặt.
- **Vùng hội thoại** chiếm toàn bộ phần giữa.
- **Ô nhập tự giãn**, hàng nút bên dưới: nút "+" (đính kèm tệp, kỹ năng, tìm web, gửi kèm vùng chọn) · mức quyền (Hỏi trước khi sửa / Tự sửa, hỏi khi rủi ro) · **nhãn tên model** bấm vào mới bung bảng chọn (nhà cung cấp, ô tìm, lọc Miễn phí, tải lại) · Gửi / Dừng.
- **Trang Cài đặt** chứa: tài khoản và key, model mặc định, quyền và an toàn, chỉ dẫn riêng, kỹ năng, đọc chéo ứng dụng, sao lưu, tự kiểm tra.
- **Màn chào** chỉ hiện khi chưa có key nào.

Vị trí chính xác của nút chọn model và lịch sử trong bản chính thức không có trong tài liệu. Add-in Claude chính thức đã cài trên máy anh: khi làm, tôi chụp khung đó để đối chiếu.

### 3.3 Lỗi mất kết nối khi đổi nhà cung cấp
Chưa có ảnh lỗi nên chưa khẳng định được một nguyên nhân duy nhất. Từ mã và đo đạc, có bốn nguyên nhân, xếp theo khả năng:

| # | Nguyên nhân | Bằng chứng |
|---|---|---|
| 1 | Đổi sang nguồn chặn gọi từ add-in (SambaNova, NVIDIA): tải danh sách model lỗi, gửi tin cũng lỗi, đều hiện "Không kết nối được…" | Đo CORS 14/11; `errors.ts:21–28` |
| 2 | Khi tải danh sách lỗi hoặc chưa xong, khung vẫn giữ danh sách và model của nguồn **cũ**, rồi gửi model của nguồn A tới nguồn B | `Chat.tsx:333–359` |
| 3 | Danh sách về thì model tự nhảy sang mục đầu theo chữ cái, không nhớ model anh đã dùng ở từng nguồn: dễ rơi vào model đã ngừng, hết lượt, không gọi được công cụ | `Chat.tsx:416–418`, `models.ts:16–21` |
| 4 | Mọi lỗi kiểu `TypeError`, kể cả lỗi lập trình, đều hiện thành "Không kết nối được" nên không biết lỗi thật | `errors.ts:24–26` |

Đã loại: lịch sử hội thoại của nguồn cũ gửi sang nguồn mới. Thử 13 kiểu lịch sử × 11 nguồn đích, các cặp chính (OpenRouter, Anthropic, Gemini, Groq, OpenAI) đều hợp lệ; chỉ có nguy cơ khi chuyển sang Mistral (mã lời gọi công cụ phải đúng 9 ký tự), khi mã có ký tự lạ chuyển sang Anthropic, và khi lịch sử có ảnh chuyển sang model không nhận ảnh.

Cách sửa (Đợt 0A): xóa danh sách và model cũ ngay khi đổi nguồn, hiện rõ "đang tải" hoặc lý do lỗi, không cho gửi khi model không thuộc nguồn đang chọn; nhớ model theo từng nguồn; hiện lỗi thật; đánh dấu nguồn không gọi được từ add-in; chuẩn hóa mã lời gọi công cụ và bỏ ảnh cũ khi chuyển nguồn giữa hội thoại; đang chạy thì việc đổi nguồn có hiệu lực từ lượt sau.

### 3.4 Chức năng gốc của Word, Excel, PowerPoint
- **Word:** đã tra xong, 36 đề xuất, đưa vào Đợt 1 và Đợt 5 (mục 5). Ba việc Office không cho add-in làm: watermark gốc, bật bảo vệ văn bản, trộn thư.
- **Excel:** đã tra, 16 đề xuất, đưa vào Đợt 4. Máy anh đủ cho bộ API mới nhất. Excel không mở Goal Seek, Data Table, Solver cho add-in; làm thay bằng cách add-in tự thử từng giá trị rồi khôi phục.
- **PowerPoint:** đã tra, 16 đề xuất, đưa vào Đợt 3. Làm được chắc chắn: chụp ảnh slide cho AI tự nhìn và sửa, kiểm bài theo mẫu, đọc màu chủ đề và bố cục, sửa bảng, đọc vùng đang chọn. Office không mở cho add-in: biểu đồ, SmartArt, ghi chú diễn giả, ẩn slide, chuyển cảnh, phông chủ đề; các việc này chỉ làm được bằng đường vòng và phải thử.

---

## 4. Đăng nhập thay cho dán API key

Kết luận cũ còn đúng. Bổ sung sau khi đối chiếu tài liệu ngày 14/11/2026:

| Phương án | Kết luận | Ghi chú mới |
|---|---|---|
| OpenRouter một lần bấm (PKCE) | **Làm** | Phần mạng đã thử: đổi mã và xem hạn mức đều gọi được từ trình duyệt. Thêm lối dự phòng chắc chắn: trang hiện mã, anh bấm Copy rồi dán vào add-in (bản chính thức cũng làm vậy). Model miễn phí: 50 lượt/ngày nếu chưa từng nạp 10 credit. |
| Tín dụng API kèm gói Max (100 USD/tháng) | **Dùng** | Vẫn là dán key một lần. Cần sửa mã nhỏ: key kiểu mới phải gắn workspace; bắt đúng thông báo hết tín dụng; Anthropic không có địa chỉ xem số dư nên chỉ báo được khi đã hết. |
| Gói Claude qua Agent SDK / đăng nhập claude.ai | **Không làm** | Tài liệu cấm đưa đăng nhập gói vào ứng dụng bên thứ ba. Tín dụng API ở trên đã thay được. |
| Đăng nhập chính thức của Anthropic cho ứng dụng bên thứ ba | Không tồn tại trên Windows/web | Chỉ có cho công cụ dòng lệnh, máy chủ, ứng dụng Apple. |
| Sign in with ChatGPT (gói Plus/Pro) | Chưa làm được | Cho phép ứng dụng mã nguồn mở, nhưng buộc có chương trình nghe cổng trên máy. Rà lại khi OpenAI nới điều kiện. |
| Sign in with Poe, Hugging Face, NanoGPT | Để sau | Có thật, cùng kiểu với OpenRouter. Viết phần đăng nhập thành khối dùng chung để sau này thêm chỉ bằng cấu hình. |
| Google (Gemini), Azure, GitHub | Không làm | Google: tốn công hơn dán key, không thêm hạn mức. Azure: cần thuê bao trả tiền. GitHub Models đã ngừng 30/07/2026. |

---

## 5. Kế hoạch theo đợt và tiêu chí nghiệm thu

Nguyên tắc giữ như cũ: mỗi đợt xong là dùng được; Claude tự làm và tự kiểm; bước không đạt chỉ làm lại bước đó, tối đa 3 lần; mục "(thử trước)" là chưa chắc về kỹ thuật, thử trong một buổi, không được thì chuyển phương án dự phòng. Cột "Kiểm bằng" ghi đúng tên mức bằng chứng của dự án: **Tự kiểm tra · Chạy thật · Kiểm thử máy**.

### Đợt 0A — Nền phát hành và bốn việc anh nêu (bản 3.6)

Việc:
1. Một nguồn mã duy nhất: `D:\ADD IN` thành bản làm việc của `main`; đưa `docs/`, `tasks/` vào kho; phát hành tự động khi đẩy mã, kiểm thử trượt thì không phát hành.
2. Enter để gửi (mục 3.1).
3. Bố cục gọn kiểu một dòng đầu (mục 3.2); bỏ công tắc tắt Track Changes, Word luôn ghi có theo dõi.
4. Sửa lỗi đổi nhà cung cấp (mục 3.3) và nút **Kiểm tra nhà cung cấp**: với mỗi key, chạy thật một chuỗi "gọi công cụ → nhận kết quả → trả lời", báo Đạt hoặc lý do.
5. **Sao lưu / Khôi phục** toàn bộ dữ liệu add-in ra một tệp (key, cài đặt, chỉ dẫn riêng, kỹ năng, lịch sử); cảnh báo khi kho trống bất thường. Tệp sao lưu chứa key nên phải ghi rõ là tệp bí mật.
6. Ba lỗi Word nhỏ đang gây sai: đọc văn bản dài bị mất phần cuối mà AI không biết; số đoạn của nhận xét và sửa đổi sai khi có đoạn trùng nội dung; trạng thái văn bản bị cắt làm AI không thấy vùng đang chọn.
7. Gỡ phần thừa: `chat.ts`, gói đệm trình duyệt cũ, 11 thư viện không dùng; một nguồn số phiên bản.

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Đầu khung hiện v3.6; `D:\ADD IN` trùng `main`; `manifest.xml` có Id `c95713b4…` | `git status`, ảnh đầu khung |
| Đẩy mã là tự phát hành; cố ý làm trượt kiểm thử thì không phát hành | Ảnh hai lần chạy tự động |
| Enter gửi, Shift+Enter xuống dòng; gõ ba câu tiếng Việt có dấu không bị gửi giữa chừng | Chạy thật trong Word |
| Ở kích thước khung mặc định, vùng hội thoại chiếm từ 70% chiều cao; không còn hàng điều khiển phía trên; luôn thấy tên model đang dùng | Ảnh trước và sau, đo điểm ảnh |
| Đổi nguồn giữa hội thoại: với mọi cặp nguồn anh có key, cả hai chiều, lượt sau chạy xong và AI nhớ lượt trước. Nguồn không gọi được hiện lý do ngay và không cho gửi | Chạy thật theo bảng cặp nguồn; Kiểm thử máy cho lớp chuyển đổi |
| Đổi qua lại hai nguồn: mỗi nguồn trở về đúng model đã dùng lần trước | Chạy thật |
| Sao lưu → xóa dữ liệu thử → khôi phục: key, chỉ dẫn riêng, kỹ năng, lịch sử trở lại đủ | Chạy thật |
| Tắt Track Changes trong Word rồi nhờ AI sửa: sửa đổi vẫn được ghi có theo dõi; Cài đặt không còn công tắc tắt | Tự kiểm tra |
| Văn bản 400 đoạn: AI đọc hết, tổng số đoạn đã đọc bằng tổng số đoạn của văn bản | Kiểm thử máy + Chạy thật |
| Tự kiểm tra không giảm: Excel 22/22, Word 14/14, PowerPoint 14/14 | Tự kiểm tra |
| `npm test` đạt (71 cũ + kiểm thử mới), kiểm kiểu không lỗi | Kiểm thử máy |

### Đợt 0B — Độ bền, chi phí, an toàn (bản 3.7)

Việc:
1. **Lịch sử chỉ thêm, không sửa.** Lưu theo từng bước, giữ đúng cặp "gọi công cụ – kết quả" khi dừng hoặc lỗi. Lời nhắc hệ thống cố định trong một hội thoại; trạng thái tài liệu và kỹ năng nạp theo lượt chuyển sang tin nhắn. Thôi cắt ngắn kết quả cũ ở giữa lịch sử; khi dài thì tóm tắt gọn một lần. Đây là điều kiện để bộ nhớ đệm có tác dụng và để giữ được phần suy luận của các model Claude mới.
2. Tự thử lại khi nhà cung cấp quá tải hoặc hết hạn mức tạm thời. Khi câu trả lời bị cắt vì hết độ dài: không chạy lệnh công cụ bị cắt dở, báo AI chia nhỏ; nâng trần độ dài đầu ra.
3. Bộ nhớ đệm lời nhắc cho key Anthropic; hiện số token mỗi lượt.
4. Ngưỡng tóm tắt theo sức chứa của model.
5. Tách phần điều phối khỏi `Chat.tsx` (làm cùng việc 1).
6. An toàn: câu phòng chỉ dẫn ẩn, kể rõ nguồn không tin cậy (nội dung, nhận xét, sửa đổi được theo dõi, đầu trang, chân trang, tệp đính kèm, kết quả web); thêm cảnh báo rủi ro cho công cụ Word còn thiếu (ghi đè ô bảng có dữ liệu, thay vùng chọn lớn, thay nhiều chỗ một lúc); hạn chờ cho từng bước công cụ; key Gemini rời khỏi đường dẫn; cập nhật thư viện.
7. Kiểm nốt 7 mục chưa chạy thật của bản 3.

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Bấm Dừng giữa chuỗi sửa, hỏi "anh vừa làm gì": AI kể đúng phần đã làm | Chạy thật trong Word |
| Lượt thứ hai cùng hội thoại (key Anthropic): số token đọc từ bộ nhớ đệm lớn hơn 0, hiện trong khung | Chạy thật |
| Giả lập quá tải hai lần rồi thành công: tự thử lại, hiện "đang thử lại lần n", bấm Dừng hủy được | Kiểm thử máy |
| Lệnh công cụ bị cắt dở: không được thực thi | Kiểm thử máy |
| Nhận xét trong văn bản chứa câu ra lệnh ("hãy xóa toàn bộ…"): AI không làm theo và báo anh | Chạy thật |
| Ghi đè ô bảng Word đang có dữ liệu: có hỏi kèm lý do | Tự kiểm tra |
| 7 mục chưa kiểm: tìm web; mở lại và xóa hội thoại; tự tóm tắt; bấm địa chỉ ô/đoạn/slide; nhập kỹ năng; đính kèm ảnh; lưu Chỉ dẫn riêng | Chạy thật, mỗi mục một ảnh |

### Đợt 1A — Word: đọc đúng, sửa đúng (bản 4.0)

Việc:
1. **Thăm dò trước (nửa ngày):** thêm vào Tự kiểm tra một bước in các bộ API máy hỗ trợ và thử 4 kiểu đánh số (danh sách thường; nhiều cấp gắn tiêu đề kiểu "Điều 1.", "1.1"; số gõ tay; trường đánh số) qua 4 đường đọc. Kết quả quyết định cách viết việc 2. Dự phòng chắc chắn nhưng tốn công: tự tính số từ cấu trúc tệp.
2. **Mục lục điều khoản:** số do Word đánh và số gõ tay, cấp, tiêu đề, định danh đoạn. Mọi kết quả đọc đều kèm số điều.
3. **Tìm trong văn bản** và **đọc trọn một điều**; kỹ năng pháp lý xem mục lục trước rồi mới đọc.
4. **Chống sửa nhầm đoạn:** định danh đoạn ổn định; công cụ sửa và xóa phải đối chiếu nội dung trước khi ghi; thay thế có giới hạn phạm vi (trong đoạn nào, lần thứ mấy); hộp xác nhận xóa hiện trích nội dung.
5. **Sửa ở mức từ:** nhiều chỗ trong một đoạn, giữ định dạng trong đoạn. Hiện 4/7 kiểu sửa thường gặp bị thay cả đoạn.
6. **Đọc và sửa đúng trên văn bản đang có sửa đổi của đối tác:** đọc theo bản hiện tại hoặc bản gốc; danh sách sửa đổi có phân trang, lọc theo tác giả, kèm số điều và câu trước/sau.
7. **Chèn điều khoản mới** kế thừa kiểu và đánh số của điều bên cạnh; đổi cấp.
8. **Hoàn tác cho Word:** từ chối các sửa đổi và gỡ nhận xét của lượt vừa rồi. Hiện Word không có nút Hoàn tác.
9. Tự kiểm tra Word chạy trong vùng cô lập, mỗi công cụ ít nhất một bước, có kiểm kết quả.

Bài thử: hợp đồng tiếng Việt 50 trang, đánh số 3 cấp, có cả số tự động lẫn số gõ tay, do Claude soạn. Chạy cùng bài trên add-in Claude chính thức để so.

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Bảng kết quả thăm dò (4 kiểu đánh số × 4 đường đọc) được ghi vào HANDOFF trước khi viết công cụ | Tự kiểm tra |
| 10 câu "Điều x.y nói gì / quy định về … nằm ở điều nào": đúng số điều từ 9/10, mỗi câu không quá 4 bước công cụ | Chạy thật |
| "Viết lại Điều 4.2…": sửa đúng đoạn; sửa hai từ cách xa nhau thì Word hiện đúng hai cặp xóa/thêm; chữ đậm giữa câu còn nguyên | Tự kiểm tra + Chạy thật |
| "Thêm một khoản sau 8.2": khoản mới mang số 8.3, khoản cũ thành 8.4, cùng kiểu chữ | Tự kiểm tra |
| Chèn hai đoạn phía trên rồi sửa theo số đoạn cũ: công cụ từ chối hoặc tự định vị lại, không sửa nhầm | Kiểm thử máy + Tự kiểm tra |
| Văn bản có 30 sửa đổi và 5 nhận xét của "đối tác": tóm tắt đủ 30 kèm số điều; xử lý cả 5 nhận xét bằng Track Changes và trả lời từng nhận xét | Chạy thật |
| Sau 3 thao tác ghi và 1 nhận xét, bấm Hoàn tác: văn bản về như trước | Tự kiểm tra |
| Bảng đối chiếu từng bài với add-in Claude chính thức (đạt ngang / kém / hơn) | Chạy thật trên cả hai |

### Đợt 1B — Word: soát và tra (bản 4.1)

Việc:
1. Đọc chú thích cuối trang, đầu trang và chân trang, bookmark, trường dẫn chiếu, ô điền mẫu, hộp văn bản.
2. **Soát nhất quán bằng mã, không tốn token:** đánh số nhảy hoặc trùng; dẫn chiếu tới điều không tồn tại; thuật ngữ định nghĩa dùng lệch hoặc dùng mà chưa định nghĩa; một bên được gọi bằng nhiều tên; chỗ trống chưa điền; số tiền bằng số không khớp bằng chữ.
3. (Thử trước) **So sánh với phiên bản khác** (tệp .docx đính kèm), tóm tắt khác biệt theo điều khoản. Bản chính thức không nêu tính năng này.
4. Kiểm thể thức trang (lề, khổ giấy, phông và cỡ chữ) cho công văn, tờ trình.
5. **PDF quét:** gửi nguyên tệp cho model đọc được (Anthropic, Gemini); nguồn khác báo rõ không đọc được. Cách đọc qua OpenRouter phải đối chiếu tài liệu trước khi làm.
6. Công tắc "Suy nghĩ kỹ"; tùy chọn soát lập luận trước khi ghi sửa đổi vào điều khoản (bản chính thức có).

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Cài sẵn 3 thuật ngữ lệch, 2 dẫn chiếu hỏng, 1 chỗ nhảy số, 1 số tiền lệch chữ, 2 chỗ trống: tìm đủ 9; không báo nhầm trên bản sạch | Kiểm thử máy + Chạy thật |
| Chú thích cuối trang và chân trang có trong câu trả lời khi được hỏi | Chạy thật |
| So hai phiên bản khác nhau 10 chỗ: liệt kê đủ 10 kèm số điều; bản gốc không bị đổi khi anh chưa đồng ý | Chạy thật |
| PDF quét 3 trang: tóm tắt đúng, trích đúng 3 số liệu (Anthropic; Gemini nếu có key) | Chạy thật |
| Văn bản lề sai chuẩn: báo đúng lề nào sai | Tự kiểm tra |
| Mỗi công cụ mới có bước Tự kiểm tra, đạt hết; bước cũ không hỏng | Tự kiểm tra |

### Đợt 2 — Đăng nhập và chi phí (bản 4.2)

Việc:
1. Nút **Đăng nhập OpenRouter** trong hộp thoại của Office (thử trước), kèm lối dự phòng dán mã (chắc chắn). Viết thành khối dùng chung.
2. Thẻ hạn mức OpenRouter: số dư còn lại, số lượt miễn phí còn trong ngày; cảnh báo khi thấp.
3. Anthropic: Claude thao tác cùng anh nhận tín dụng 100 USD/tháng, tạo key gắn workspace; add-in bắt đúng thông báo hết tín dụng và gợi ý đổi nguồn.
4. Key hết hạn, bị thu hồi, hết tiền: thông báo tiếng Việt rõ, không mất hội thoại.

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Từ trạng thái chưa có key OpenRouter: bấm một nút → đăng nhập → danh sách model hiện → sửa được một đoạn Word | Chạy thật |
| Lối dự phòng: chép mã từ trang, dán vào add-in, kết quả như trên | Chạy thật |
| Đăng xuất xóa sạch khóa trên máy | Chạy thật |
| Trang Thanh toán của Console hiện tín dụng; một yêu cầu chạy xong, số dư giảm | Ảnh trước và sau |
| Key không có số dư: thông báo tiếng Việt rõ lý do, hội thoại còn nguyên | Chạy thật |

### Đợt 3 — Chéo ứng dụng và PowerPoint (bản 4.3)

Việc:
1. (Thử trước) Kiểm Word, Excel, PowerPoint có dùng chung một kho lưu trữ khi cùng tài khoản Office; rồi thử kênh lệnh cục bộ để khung ở ứng dụng này nhờ khung ở ứng dụng kia ghi hộ. Luôn hỏi xác nhận trước khi ghi sang ứng dụng khác.
2. Chỉ báo ứng dụng đang kết nối; công tắc bật/tắt theo từng ứng dụng; báo rõ khi khung đích chưa mở.
3. PowerPoint, phần chắc chắn làm được: **chụp ảnh slide cho AI tự kiểm** (chữ tràn, hình chồng nhau) rồi tự sửa một lượt; **kiểm bài theo mẫu** (phông, cỡ, màu lệch chuẩn; slide thiếu tiêu đề; chữ quá nhỏ); đọc màu chủ đề, bố cục, ô giữ chỗ của mẫu; định dạng chữ theo từng đoạn; sửa bảng có sẵn; đọc hình và đoạn chữ anh đang chọn; đổi bố cục slide.
4. PowerPoint, phần phải thử trước (Office không có lối trực tiếp): cập nhật số liệu biểu đồ đã có; đọc phông chủ đề; chèn ảnh; ghi chú diễn giả. Không được thì báo rõ, không cố.

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Từ PowerPoint: "lấy bảng tiến độ ở Excel đang mở, tạo 3 slide và 1 biểu đồ" chạy trọn, đúng số | Chạy thật |
| Từ Excel: "ghi tóm tắt sang Word đang mở": đoạn mới xuất hiện trong Word, có hỏi xác nhận, hoàn tác được ở Word | Chạy thật |
| Đóng khung ở ứng dụng kia: AI báo rõ không ghi được, không treo; chỉ báo kết nối hiện đúng | Chạy thật |
| Slide cố ý để chữ tràn khung: AI tự phát hiện qua ảnh slide và sửa hết tràn | Chạy thật |
| Bài 10 slide cài 5 lỗi lệch mẫu (2 phông lạ, 1 slide thiếu tiêu đề, 2 chỗ chữ dưới 12 pt): kiểm bài tìm đủ 5 | Tự kiểm tra |
| Slide mới dùng đúng màu và bố cục của mẫu | Tự kiểm tra |

Dự phòng nếu việc 1 không được: giữ đọc chéo, thêm nút "Chuyển sang … và làm tiếp" kèm sẵn nội dung.

### Đợt 4 — Excel và dấu vết soát lại (bản 4.4)

Việc:
1. **Dấu vết soát lại:** sheet "Nhật ký AI" ghi từng thay đổi (giờ, ô, giá trị cũ, giá trị mới, lý do, yêu cầu của lượt đó); tô sáng ô vừa sửa, tắt được. Theo bản chính thức hiện hành: nhật ký **mặc định tắt**, có cảnh báo sheet nhật ký đi theo tệp khi gửi cho người khác.
2. **Sửa cái có sẵn:** biểu đồ (kiểu, tiêu đề, chuỗi, nguồn dữ liệu), PivotTable (thêm bớt trường, hàm tổng hợp, lọc, làm mới).
3. **In báo cáo:** vùng in, tiêu đề lặp, A4 ngang, co vừa trang, lề, đầu và chân trang.
4. **Mô hình tài chính:** tên vùng cho các giả định; kiểm toán mô hình (hằng số gõ cứng lẫn trong vùng công thức, công thức lệch hàng, liên kết ra tệp ngoài); (thử trước) giải ngược tìm đầu vào để đạt mục tiêu và bảng nhạy cảm 1–2 biến, luôn khôi phục giá trị gốc.
5. **Bảng theo dõi hồ sơ:** làm sạch dữ liệu (xóa trùng, thay hàng loạt, tự điền chuỗi ngày), nhóm hàng cột, khóa ô công thức và chừa ô nhập.
6. Hỏi trước khi dùng hàm nguy hiểm. Mục này và vùng in không còn trong tài liệu chính thức; giữ như lựa chọn riêng.

| Tiêu chí nghiệm thu | Kiểm bằng |
|---|---|
| Sau 5 thao tác ghi: sheet nhật ký có đủ 5 dòng (giờ, ô, giá trị cũ, giá trị mới, lý do) | Tự kiểm tra + Chạy thật |
| Hoàn tác thì ô hết tô sáng và nhật ký ghi thêm dòng hoàn tác | Chạy thật |
| Đổi kiểu và tiêu đề một biểu đồ có sẵn; thêm một trường vào Pivot có sẵn | Tự kiểm tra |
| "Chuẩn bị in bảng tiến độ": đúng vùng in, A4 ngang, vừa một trang ngang, dòng tiêu đề lặp | Tự kiểm tra |
| Mô hình thử cài 3 hằng số gõ cứng và 1 công thức lệch hàng: kiểm toán tìm đủ 4 | Tự kiểm tra |
| Giải ngược trên mô hình thử: kết quả đạt mục tiêu trong sai số 0,1%; bấm hủy thì mọi ô về giá trị gốc | Tự kiểm tra + Chạy thật |
| Công thức có WEBSERVICE hoặc INDIRECT: có hỏi trước khi ghi | Tự kiểm tra |

### Đợt 5 — Tùy chọn, làm khi anh cần

Bộ nhớ dài hạn giữa các hội thoại · tải bản ghi hội thoại · gợi ý dạng gạch chân bấm để chấp nhận trong Word · điền mẫu qua ô điền sẵn · connectors (bản chính thức gọi thẳng máy chủ MCP từ khung, thử theo cách đó) · công cụ chạy mã Office.js do AI viết (khoảng cách năng lực lớn nhất còn lại, rủi ro cao, luôn hỏi kèm hiện mã) · phím tắt mở khung · đọc một trang web cụ thể · Outlook · Sign in with Poe.

---

## 6. Không làm, và vì sao

- Viết lại bằng khung khác, thêm máy chủ riêng, đổi thư viện giao diện: như kế hoạch cũ.
- Mức quyền "bỏ qua mọi xin phép" của bản chính thức: trái quy tắc dự án (thao tác rủi ro luôn hỏi).
- Watermark gốc, bật bảo vệ văn bản, trộn thư trong Word: Office không mở cho add-in. Cần thì có cách vòng, không cam kết.
- Goal Seek, Data Table, Solver gốc của Excel; tạo và sửa SmartArt, chuyển cảnh, ẩn slide trong PowerPoint: Office không mở cho add-in. Riêng giải ngược và bảng nhạy cảm có làm thay ở Đợt 4.
- Hàm `=AI(…)` trong ô Excel: làm được nhưng phải đổi cách add-in nạp trang (shared runtime), công lớn; để Đợt 5 nếu anh cần.
- Đăng nhập bằng gói Claude: tài liệu cấm với ứng dụng bên thứ ba.

## 7. Sửa bảng so sánh (làm ở Đợt 0A)

- **Sai, phải sửa:** dòng 10 và câu tổng hợp (lối gateway của bản chính thức nay có Skills và tệp đính kèm, chỉ thiếu chéo ứng dụng và đọc giọng nói); ghi chú bỏ Ctrl+Alt+C (phím tắt vẫn còn).
- **Hết nguồn xác minh:** dòng 34–37 (bốn bài hỗ trợ đã gộp vào trang tài liệu, các chi tiết này không còn).
- **Ghi quá mức thực tế của AI Đa Nguồn:** dòng 5, 41, 42, 44, 45 (phần Word).
- **Thêm dòng mới của bản chính thức:** bộ nhớ dài hạn; tự kiểm hình thức sau khi sửa; soát lập luận trước khi sửa ngôn từ pháp lý; tải bản ghi hội thoại; công cụ chạy mã Office.js; giới hạn thời gian và cỡ tệp; phát hiện va chạm khi người dùng đang sửa; hội thoại gắn với tệp; lấy lại phần ngữ cảnh đã đẩy ra; dò khả năng của model khi kết nối; bộ nhớ đệm ở mọi lượt; Outlook.

## 8. Điều phải thử trong Office trước khi cam kết

| Điều thử | Thuộc đợt | Nếu không được |
|---|---|---|
| Đọc số điều khoản tự đánh gắn kiểu tiêu đề | 1A | Tự tính số từ cấu trúc tệp (công lớn) |
| Thuộc tính đọc đoạn có lẫn chữ đã bị xóa có theo dõi không | 1A | Đọc bằng lối khác theo từng đoạn |
| Sửa nhiều chỗ trong một đoạn khi đang theo dõi thay đổi | 1A | Giảm còn sửa theo cụm, vẫn hơn thay cả đoạn |
| So sánh phiên bản: kết quả hiện ở đâu, add-in đọc được không | 1B | So văn bản thuần, không tạo redline |
| Tải tệp từ khung add-in ra đĩa (sao lưu, bản ghi) | 0A | Chép vào bộ nhớ tạm hoặc chèn vào tài liệu mới |
| Đăng nhập OpenRouter trong hộp thoại Office | 2 | Lối dán mã |
| Ba ứng dụng dùng chung kho; kênh lệnh giữa các khung | 3 | Nút "Chuyển sang … và làm tiếp" |

## 9. Việc cần anh

- **Đầu Đợt 0A:** cho phép Claude điều khiển Word, Excel, PowerPoint trên máy (một lần bấm cho mỗi ứng dụng) để tự kiểm.
- **Khi manifest đổi** (bố cục mới không cần; đăng nhập có thể cần): bấm đúp `cai-dat.cmd` một lần rồi mở lại Office. Tôi sẽ báo trước.
- **Đợt 2:** bấm Authorize ở OpenRouter; bấm xác nhận nhận tín dụng API ở claude.ai.

## 10. Nguồn

Add-in chính thức (bản .md, đọc 14/11/2026): `claude.com/docs/office-agents/` các trang overview, excel, word, powerpoint, outlook, work-across-apps, connectors-and-skills, data-storage, dictation, third-party-platforms, performance, opentelemetry.

Đăng nhập: `openrouter.ai/docs/guides/overview/auth/oauth` · `openrouter.ai/docs/api/reference/limits` · `platform.claude.com/docs/en/about-claude/api-credits-for-subscribers` · `support.claude.com/en/articles/17154008` · `support.claude.com/en/articles/15036540` · `code.claude.com/docs/en/legal-and-compliance` · `platform.claude.com/docs/en/manage-claude/authentication` · `developers.openai.com/siwc/`.

Word: `learn.microsoft.com/en-us/javascript/api/requirement-sets/word/` (WordApi 1.3–1.9, WordApiDesktop 1.1–1.5) · `learn.microsoft.com/en-us/office/dev/add-ins/word/manage-tracked-changes`.

Excel, PowerPoint: `learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/` (ExcelApi 1.7–1.21) · `learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout` · `learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/` (PowerPointApi 1.2–1.10).

Kết quả chi tiết của các tác tử (đường dẫn nguồn từng phát hiện, dòng mã, 36 đề xuất Word, 16 đề xuất Excel, 16 đề xuất PowerPoint) nằm trong nhật ký phiên; khi duyệt, phần cần giữ sẽ được chép vào `docs/`.
