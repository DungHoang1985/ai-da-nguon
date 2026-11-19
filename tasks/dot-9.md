# Đợt 9 – Hai chế độ Miễn phí / Trả phí và xoay vòng nguồn miễn phí

Ngày lập: 19/11/2026. Nhánh: `claude/dot-9` (tách từ `main@8f4a4f7`). Căn cứ: mục 12 của `tasks/ban-giao-phan-phoi-cong-ty-va-che-do-web.md` (anh duyệt khoảng 00:00 ngày 19/11, sửa lúc 00:03 để đưa Cohere vào chế độ Miễn phí) và năm ý anh đưa ra sáng 19/11.

**Trạng thái: CHỜ DUYỆT bản sửa.** Anh yêu cầu xem xét kỹ nội dung đã duyệt lúc 0 giờ. Lượt rà mã dưới đây thấy 9 chỗ kế hoạch gốc chưa tính đến; kế hoạch và tiêu chí đã sửa theo các chỗ đó. Chưa sửa mã.

Nhãn: **[Đã đọc mã]** = đối chiếu tệp trên `main@8f4a4f7`. **[Chưa kiểm]** = phải kiểm tài liệu chính thức hoặc đo thật khi làm.

---

## 1. Ý anh ngày 19/11/2026 (sáng)

| # | Ý anh | Đưa vào kế hoạch |
|---|---|---|
| 1 | Sửa lại chỗ làm cho đúng | Đã làm: tệp bàn giao và `HANDOFF.md` đã commit lên `main` (`8f4a4f7`, chưa đẩy lên GitHub); Đợt 9 làm trên nhánh `claude/dot-9` tách từ `main`, không làm trên `gh-pages`. |
| 2 | Xem xét kỹ nội dung duyệt lúc 0 giờ | Mục 2 dưới đây. |
| 3 | Im lặng 150 giây là không phản hồi: xoay ngay sang nguồn khác | Bước 4: im lặng → chuyển nguồn kế tiếp, không gọi lại nguồn đó trong lượt. |
| 4 | Lỗi 413 (yêu cầu quá lớn) cũng chuyển nguồn | Bước 4; gộp cả lỗi vượt sức chứa ngữ cảnh (cùng bản chất). |
| 5 | Có nút chọn chế độ Miễn phí / Trả phí để người dùng chọn khi bắt đầu | Bước 2: thẻ chọn chế độ lúc mở khung lần đầu, kể cả người đang dùng bản cũ; công tắc luôn thấy ở bảng chọn model. Thay cho ý "mặc định Miễn phí" của mục 12. |

---

## 2. Rà kế hoạch mục 12: những chỗ chưa tính đến

| # | Phát hiện | Căn cứ | Hệ quả nếu giữ nguyên | Sửa |
|---|---|---|---|---|
| R1 | **"Miễn phí" thuộc về key, không thuộc về nguồn.** Key Gemini của dự án Google đã bật thanh toán bị tính tiền mọi lần gọi; Groq gói Developer và Cohere key sản xuất cũng vậy. Add-in không đọc được key thuộc gói nào. Chỉ model `:free` của OpenRouter chắc chắn không tính tiền. | [Chưa kiểm] trang giá Gemini, Groq, Cohere ngày làm | Tiêu chí T4 cũ ("không bao giờ gọi nguồn trả phí") không chứng minh được bằng mã; người dùng tưởng miễn phí mà bị trừ tiền. | Chế độ Miễn phí chỉ hiện model có bậc miễn phí (R2). Khi nhập key Gemini, Groq hoặc Cohere ở chế độ này, hiện dòng: "Chỉ miễn phí khi key thuộc gói miễn phí (dự án / tài khoản chưa bật thanh toán). Đã bật thanh toán thì mỗi lần gọi bị tính tiền." T4 viết lại theo đúng phần mã kiểm được. |
| R2 | **Không phải model Gemini nào cũng có bậc miễn phí** (dòng Pro thường không có). Danh mục model của Google không ghi giá. | [Đã đọc mã] `catalog.ts` lọc "Miễn phí" chỉ áp cho OpenRouter (`m.free`) | Chế độ Miễn phí vẫn hiện model Pro, gọi vào thì bị 429 hoặc bị tính tiền. | Danh sách cho phép theo nguồn trong chế độ Miễn phí: Gemini chỉ Flash, Flash-Lite, Gemma; Groq tất cả; OpenRouter chỉ `:free`; Cohere tất cả (key dùng thử). Kiểm lại với trang giá chính thức ngày làm. |
| R3 | **Chuyển sang Gemini 3 giữa lượt có thể bị từ chối.** Lệnh gọi công cụ do nguồn khác sinh ra được dựng lại thành `functionCall` không có `thoughtSignature`; Gemini 3 đòi chữ ký này cho lệnh gọi công cụ trong lượt hiện tại. | [Đã đọc mã] `llm.ts` dòng 213–220. [Chưa kiểm] tài liệu Gemini về chuyển lịch sử từ model khác | Xoay Groq → Gemini giữa lượt bị lỗi 400, vòng xoay đứt đúng lúc cần. | Đối chiếu tài liệu Gemini (có cách khai chữ ký thay thế cho lịch sử của model khác); thêm kiểm thử dựng yêu cầu. Không có cách hợp lệ thì khi xoay sang Gemini, nén các bước công cụ trước thành chữ. |
| R4 | **Sức chứa của các nguồn chênh nhau nhiều.** Gemini khoảng 1 triệu token; Groq nhỏ hơn nhiều và gói free giới hạn token mỗi phút (413 đã gặp ở Đợt 4); model `:free` của OpenRouter mỗi cái một khác. | [Đã đọc mã] `errors.ts` lý do `too_large`; `compaction.ts` tính ngưỡng tóm tắt theo model đang dùng | Một lượt Word 11–35 nghìn token chạy được trên Gemini, sang Groq có thể quá cỡ; ngưỡng tóm tắt tính theo Gemini thì lịch sử quá dài cho nguồn sau. | 413 và lỗi vượt ngữ cảnh → chuyển nguồn (ý 4). Nguồn không đủ sức chứa cho lượt hiện tại thì bỏ qua luôn, không gọi thử. |
| R5 | **Mỗi nguồn trong vòng cần một model cụ thể.** Kế hoạch gốc chỉ nêu thứ tự nguồn. Model khác nhau về việc nhận ảnh, gọi công cụ, sức chứa. | [Đã đọc mã] `loop.ts` nhận một `provider`, `model`, `vision`, `maxOutput` cho cả lượt | Không rõ gọi model nào; ảnh slide hoặc trang PDF quét gửi tới model không nhận ảnh là bị từ chối cả lượt (quy tắc Đợt 5). | Vòng xoay là danh sách (nguồn, model). Mặc định lấy model đã nhớ của nguồn, chưa có thì lấy model ưu tiên gọi được công cụ (`PREFER`). Mỗi lần xoay cập nhật `vision`, sức chứa, trần đầu ra theo model mới. |
| R6 | **Chữ đã hiện rồi mới lỗi.** Hiện nay đã có chữ hiện ra thì không tự thử lại (tránh lặp chữ). | [Đã đọc mã] `callModelStep`, biến `emitted` | Lượt dừng giữa chừng dù còn nguồn khác. | Vẫn xoay: phần chữ dở giữ trên khung kèm dòng "Đổi sang <nguồn>: trả lời lại bước này"; phần dở không vào lịch sử gửi model; nguồn mới làm lại bước đó. |
| R7 | **Kế hoạch gốc không nói rõ có thử lại cùng nguồn trước khi xoay không.** Hiện nay 429 theo phút / 5xx tự thử lại 3 lần, chờ 2, 6, 15 giây. | [Đã đọc mã] `RETRY_DELAYS_MS` | Mất tới 23 giây trước khi xoay. | Chế độ Miễn phí: xoay ngay; nguồn vừa lỗi tạm thời được "nghỉ" 60 giây hoặc theo `retry-after`, rồi mới vào lại vòng. Chế độ Trả phí giữ cách thử lại cũ. |
| R8 | **Mốc đặt lại hạn mức ngày khác nhau theo nguồn.** | [Chưa kiểm] Gemini đặt lại lúc nửa đêm giờ Thái Bình Dương (khoảng 15:00 giờ Việt Nam); OpenRouter theo ngày UTC (7:00 giờ Việt Nam); Groq, Cohere (theo tháng) phải kiểm | Đánh dấu "hết hôm nay" sai giờ: bỏ phí nguồn còn dùng được, hoặc gọi lại nguồn còn hết. | Bảng mốc đặt lại theo nguồn, có căn cứ; không rõ thì nghỉ 24 giờ kể từ lúc lỗi. Lưu cục bộ để mở lại khung vẫn nhớ. |
| R9 | **Nguồn phải có key mới vào được vòng.** | — | Vòng toàn nguồn không có key, xoay ra lỗi "chưa nhập key". | Vòng chỉ gồm nguồn đã có key. Ít hơn 2 nguồn thì hiện dòng gợi ý nhập thêm key (kèm nút tới trang lấy key). |

Các điểm của mục 12 giữ nguyên: thứ tự mặc định **Gemini → Groq → OpenRouter `:free` → Cohere**, sửa được trong Cài đặt; chế độ Trả phí không xoay vòng; **không bao giờ tự chuyển từ Miễn phí sang Trả phí**; hết mọi nguồn thì dừng và báo; nguồn tự khai do người dùng chọn chế độ khi khai; Mistral xếp Trả phí; Cohere ghi chú "Key dùng thử: miễn phí, giới hạn tốc độ, Cohere không cho dùng vào mục đích thương mại."

Xếp lại các nguồn "chưa xếp": Cerebras → Trả phí (hết bậc free thường trực 16/7/2026); SambaNova, NVIDIA vẫn bị chặn CORS → ẩn ở cả hai chế độ như hiện nay.

---

## 3. Phân loại nguồn

| Chế độ | Nguồn (model) |
|---|---|
| Miễn phí | Gemini (Flash, Flash-Lite, Gemma) · Groq · OpenRouter (chỉ `:free`) · Cohere (key dùng thử) |
| Trả phí | Anthropic · OpenRouter (model trả phí) · OpenRouter·Claude · DeepSeek (3 kênh) · MiniMax · Kimi · OpenAI · Mistral · Cerebras · Gemini, Groq, Cohere khi key đã bật thanh toán (người dùng tự chọn) |
| Ẩn | SambaNova, NVIDIA (chặn CORS) |
| Nguồn tự khai | Người dùng chọn chế độ khi khai |

---

## 4. Các bước

1. **Lõi (`providers.ts`, `catalog.ts`):** trường `tier` ("free" / "paid" / "both") cho từng nguồn, danh sách model cho phép ở chế độ Miễn phí; hàm thuần `providersForMode`, `modelsForMode`.
2. **Chọn chế độ (ý 5):** thẻ "Chọn chế độ" hiện lúc mở khung khi chưa chọn (cả người dùng bản cũ, không tự đổi nguồn họ đang dùng): hai nút **Miễn phí** (kèm câu: dùng key gói free, tự đổi nguồn khi hết lượt, dữ liệu có thể được nhà cung cấp dùng để cải thiện dịch vụ) và **Trả phí** (kèm câu: tính tiền theo key của anh, không tự đổi nguồn). Công tắc chế độ luôn thấy ở đầu bảng chọn model; đổi được trong Cài đặt. Nhãn model ghi chế độ đang bật.
3. **Lọc (`ModelPicker`, Cài đặt):** bộ chọn nguồn và model chỉ hiện nguồn, model của chế độ đang bật; dòng cảnh báo R1 khi nhập key Gemini, Groq, Cohere ở chế độ Miễn phí.
4. **Xoay vòng (lõi, chỉ chế độ Miễn phí), hàm thuần `rotation.ts` + nối vào `loop.ts` / `llm.ts`:**
   - Chuyển nguồn kế tiếp ngay khi gặp: im lặng 150 giây (ý 3), 429 theo phút, 5xx / 503 / 529, lỗi mạng, 413 và lỗi vượt ngữ cảnh (ý 4), lỗi nguồn không gọi được công cụ.
   - 429 theo ngày → nguồn nghỉ đến mốc đặt lại (R8). Lỗi tạm thời → nghỉ 60 giây hoặc theo `retry-after` (R7).
   - Lỗi key hoặc tiền (`reason`) → bỏ nguồn đó trong lượt, báo trên khung, vẫn xoay tiếp.
   - Gọi lại bằng `plain` (bỏ khối riêng của nguồn trước); xử lý chữ ký Gemini (R3); cập nhật `vision`, sức chứa theo model mới (R4, R5); chữ dở (R6).
   - Khung hiện dòng "Đang dùng <nguồn · model>" và mỗi lần đổi: "Đổi sang <nguồn> vì <lý do>".
   - Hết nguồn → dừng, báo từng nguồn hỏng vì sao và khi nào dùng lại được; không bao giờ gọi nguồn Trả phí.
5. **Kiểm thử máy (`npm test`):** lọc theo chế độ; xoay vòng với lỗi giả lập (im lặng, 429 phút, 429 ngày, 503, 413, vượt ngữ cảnh, lỗi key); đổi Groq → Gemini có lệnh gọi công cụ trong lịch sử; chữ dở; không có nguồn trả phí nào trong vòng.
6. **Thử thật trong Office (desktop ẩn, bản cục bộ `cdp.mjs serve`):** chế độ Trả phí 1 bài Word bằng key Claude đã nhập sẵn; chế độ Miễn phí 1 bài Word với các key free đang có, ép xoay bằng cách cho nguồn đầu im lặng (bản thử, không vào đường chính). Gemini free hết hạn mức thì ghi "Chưa kiểm được".
7. **Soát mã độc lập** phần xoay vòng (đụng lõi tác tử), sửa theo kết quả rồi mới phát hành bản **4.9.0**.

---

## 5. Tiêu chí nghiệm thu

| # | Tiêu chí | Bằng chứng |
|---|---|---|
| T1 | Chế độ Miễn phí chỉ hiện nguồn, model có bậc miễn phí (theo bảng mục 3); Trả phí chỉ hiện nguồn, model thu phí | Kiểm thử máy + ảnh khung hai chế độ |
| T2 | Lần mở khung đầu tiên (cả khi nâng từ 4.8.0 có key Claude) hiện thẻ chọn chế độ; chọn xong không hiện lại; nguồn đang dùng không bị đổi ngầm | Kiểm thử máy + ảnh khung |
| T3 | Im lặng, 429 theo phút, 503, 413, vượt ngữ cảnh ở nguồn free → chuyển nguồn kế tiếp ngay, lượt làm xong | Kiểm thử giả lập lỗi + nhật ký một lượt thật |
| T4 | Chế độ Miễn phí không bao giờ gọi nguồn hoặc model ngoài danh sách Miễn phí, kể cả khi mọi nguồn free hỏng; cảnh báo R1 hiện khi nhập key Gemini / Groq / Cohere | Kiểm thử máy + ảnh khung |
| T5 | 429 theo ngày → nguồn bị bỏ qua đến mốc đặt lại, kể cả sau khi mở lại khung | Kiểm thử máy |
| T6 | Đổi nguồn giữa lượt có lệnh gọi công cụ (Groq → Gemini, Gemini → OpenRouter) gửi được yêu cầu hợp lệ | Kiểm thử máy dựng yêu cầu + 1 lần gọi thật nếu còn hạn mức |
| T7 | Chế độ Trả phí chạy đúng 1 bài Word bằng key Claude có sẵn, không xoay vòng | JSON trong `tasks/dot-9-bang-chung/` |
| T8 | Cohere gọi được công cụ qua `compatibility/v1` | 1 lần gọi thật, hoặc ghi "Chưa kiểm được" kèm lý do |
| T9 | Kiểm thử cũ vẫn xanh (609+), kiểm kiểu sạch; Tự kiểm tra Word / Excel / PowerPoint không lùi; không lộ key trong bằng chứng | `npm test`, `tsc`, JSON Tự kiểm tra, rà bằng chứng |

---

## 6. Điểm cần anh chốt trước khi viết mã

1. Hạn chờ im lặng ở chế độ Miễn phí: giữ 150 giây hay hạ xuống (ví dụ 60 giây) để xoay sớm hơn. Hạ thì có thể cắt nhầm lúc Gemini 3 đang "Suy nghĩ kỹ" lâu.
2. Key Gemini / Groq / Cohere có thể bị tính tiền (R1): chỉ hiện cảnh báo, hay bắt người dùng tích xác nhận "key gói miễn phí" trước khi nguồn vào vòng.
3. Lỗi 429 theo phút: xoay ngay, hay chờ một lần khi nhà cung cấp hẹn ≤ 10 giây rồi mới xoay.

## 7. Việc buộc phải có anh ở máy

Không có: các key cần dùng đã lưu sẵn trong add-in. Bài thử Trả phí tốn một ít tín dụng key Claude (một bài Word, ước chừng dưới 0,5 USD [Chưa kiểm]).
