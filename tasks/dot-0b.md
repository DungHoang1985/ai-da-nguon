# Đợt 0B — Độ bền, chi phí, an toàn (v3.6 → v3.7)

Lập 15/11/2026 cuối phiên Đợt 0A (bản nháp, rà lại trước khi làm). Tiêu chí nghiệm thu: bảng Đợt 0B trong `tasks/ke-hoach-v4.md` mục 5.

## Việc còn lại của Đợt 0A — kết quả kiểm trong Office bản 3.6 và 3.6.1 (15/11/2026, trên máy)

Bằng chứng: `tasks/dot-0a-bang-chung/08`–`22`. Kiểm trên Word (Document1 trống), Excel (Book1), PowerPoint (Presentation1), Office 16.0.20527, khung add-in tải từ GitHub Pages.

| Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|
| Đầu khung hiện v3.6; bố cục một dòng đầu; vùng hội thoại ≥ 70% chiều cao | **Đạt** — khung 595 điểm ảnh cao, vùng hội thoại khoảng 470 (79%); không còn hàng điều khiển phía trên; nhãn model luôn hiện cạnh ô nhập | Ảnh 09 |
| Enter gửi, Shift+Enter xuống dòng | **Đạt** | Chạy thật trong Word (tin hai dòng gửi bằng Enter, AI trả lời) |
| Gõ ba câu tiếng Việt có dấu bằng Unikey (Telex) không bị gửi giữa chừng | **Đạt** — "Tôi đang thử gõ / Hợp đồng này dài / Điều khoản bảo mật" còn nguyên trong ô nhập, Enter mới gửi | Ảnh 10 |
| Như trên với bộ gõ Telex của Windows | **Không kiểm nữa** (anh quyết định 15/11/2026): anh dùng Unikey và đã tự dùng thử, gõ tiếng Việt rồi Enter vẫn ổn. Máy không cài bộ gõ Telex của Windows. Phần chặn khi đang ghép chữ (`isComposing`) vẫn có Kiểm thử máy | Lời anh trong phiên 15/11 buổi chiều; `Get-WinUserLanguageList` |
| Đổi qua lại hai nguồn: mỗi nguồn về đúng model đã dùng | **Đạt** — Anthropic ↔ Groq: Anthropic về Claude Haiku 5.5, Groq về GPT OSS 20B (không phải mục đầu danh sách) | Ảnh 11, 12 |
| Kiểm tra nhà cung cấp: OpenAI báo đúng nguyên nhân, không còn "Không kết nối được" | **Đạt** — "OpenAI trả về lỗi nhưng không cho add-in đọc chi tiết. Thường do: key sai hoặc hết hạn, tài khoản hết tiền, hoặc tài khoản chưa được dùng model gpt-5.5…". Từng nguồn: đạt 4/8 (Groq, Anthropic, Cerebras, Cohere); Gemini, DeepSeek hết tiền; OpenRouter trượt vì model đang nhớ | Tệp 08 |
| Sao lưu → xóa dữ liệu thử → khôi phục gộp | **Đạt** — tệp có 8 key, 1 chỉ dẫn riêng, 6 hội thoại, 1 lệnh; xóa chỉ dẫn thử, lệnh thử, hội thoại thử rồi "Gộp vào dữ liệu hiện có": cả ba trở lại, 8 key còn nguyên. Kỹ năng nhập thêm: chưa kiểm (máy chưa có kỹ năng nào nhập thêm) | Ảnh 13, 14, 15 |
| Tải tệp từ khung add-in ra đĩa | **Đạt** — tệp `AI-Da-Nguon-sao-luu-…-BI-MAT.json` nằm trong thư mục Tải xuống; không cần lối "Chép vào bộ nhớ tạm" | Ảnh 13 |
| Tự kiểm tra Word 17 bước | **Đạt 17/17 ở bản 3.6.1** (13:02). Bản 3.6 ra 16/17: bước "Tắt Track Changes rồi AI sửa: vẫn ghi có theo dõi" báo "2 → 2" vì bước kiểm sửa ngay trong đoạn AI vừa chèn (đang là sửa đổi chờ duyệt) nên Word gộp, số sửa đổi không tăng. Sản phẩm vẫn đúng: chạy thật, Track Changes tắt, AI chèn rồi thay chữ đều ra sửa đổi có theo dõi. Sau khi sửa bước kiểm: "sửa đổi vẫn được ghi (2 → 3), trả lại chế độ Off" | Ảnh 16, 17, 21 |
| Tự kiểm tra Excel | **Đạt 22/22** | Ảnh 18 |
| Tự kiểm tra PowerPoint | **Đạt 14/14 ở lần chạy thứ hai**; lần đầu trên bài mới tinh 13/14: bước "Biểu đồ gốc" báo mẫu bị thêm bố cục (11 → 12) | Ảnh 19, 20 |
| Sửa `docs/so-sanh-claude-chinh-thuc.md` theo mục 7 | **Đạt** — 67 dòng, đã đếm lại cột trạng thái | `git diff` |

### Lỗi tìm ra khi kiểm, và việc đã làm

1. **Bước Tự kiểm tra Word đo sai** (không phải lỗi sản phẩm). Đã sửa `src/office/selftest.ts`: chèn một đoạn thường khi đang tắt theo dõi rồi mới nhờ công cụ sửa. **Đã phát hành 3.6.1 và chạy lại trong Word: 17/17** (ảnh 21).
2. **`word_replace_text` không thay gì khi model truyền `max_replacements` là -1 hoặc số âm** mà vẫn báo "Đã thay -1 chỗ" (gặp với Groq · GPT OSS 20B). Đã sửa (`replaceLimit` trong `src/office/util.ts`), có 2 kiểm thử máy; `npm test` 120/120, kiểm kiểu sạch. **Đã chạy thật ở bản 3.6.1** (Anthropic · Claude Haiku 5.5): tham số `{"find":"xanh","replace":"tim","max_replacements":-1}` → kết quả `{"ok":true,"found":1,"replaced":1}`, văn bản đổi đúng (ảnh 22).
3. **PowerPoint: lần chèn biểu đồ đầu tiên trên bài mới tinh làm mẫu thêm một bố cục** dù mã đã gắn kiểu bố cục. Kết quả "14/14" của các bản trước nhiều khả năng đo trên bài đã có sẵn bố cục thừa. Chưa sửa → đưa vào Đợt 3 (hoặc sớm hơn nếu anh hay dùng biểu đồ).
4. **OpenRouter: model đang nhớ `thinkingmachines/inkling-2-small:free` bị OpenRouter chỉ cho dùng trong "agentic harnesses"**; add-in báo nhầm thành "Key không đúng hoặc không có quyền". Chưa sửa → gộp vào việc 3 dưới đây.
5. Nhỏ: bấm Gửi khi bảng Lịch sử đang mở thì bảng vẫn che hội thoại, phải bấm Đóng. Chưa sửa.

### Còn lại, chuyển sang Đợt 0B

1. ~~Phát hành hai sửa đổi ở trên rồi chạy lại Tự kiểm tra Word~~ — xong ở bản 3.6.1 (commit `b0fe2aa`).
2. ~~Kiểm Enter với bộ gõ Telex của Windows~~ — anh bỏ mục này 15/11/2026: anh dùng Unikey, đã tự thử, ổn.
3. Chọn model mặc định: OpenAI `gpt-5.5` có thể cần xác minh tổ chức; OpenRouter tránh model bị giới hạn nơi dùng và báo đúng lý do khi gặp.
4. Tệp sao lưu thử còn trong thư mục Tải xuống (chứa 8 key thật và ba mục dữ liệu thử): giữ làm bản sao lưu hoặc xóa, tùy anh.

## Việc của Đợt 0B (theo kế hoạch mục 5)

Lịch sử chỉ thêm, không sửa · tự thử lại khi 429/5xx, không chạy lệnh công cụ bị cắt dở · bộ nhớ đệm lời nhắc Anthropic, hiện số token · ngưỡng tóm tắt theo sức chứa model · tách điều phối khỏi `Chat.tsx` (hook `useCatalog` đã tách phần danh mục) · an toàn (chỉ dẫn ẩn, cảnh báo rủi ro Word còn thiếu, key Gemini rời khỏi đường dẫn) · kiểm nốt 7 mục chưa chạy thật. Trước khi làm: rà nốt mã lõi tác tử (`loop.ts`, `llm.ts`, `compaction.ts`).

## Kết quả Đợt 0B — bản 3.7.0 (15/11/2026 buổi chiều, phiên trên máy)

Phát hành: commit `0c570a7` trên `main`, gh-pages có "Bản chạy thật v3.7.0 (main@0c570a7)", GitHub Pages phục vụ `taskpane.1414ffa9.js` chứa số bản 3.7.0 (đo bằng curl lúc phát hành).

**Chưa kiểm được trong Office ở phiên này:** yêu cầu quyền điều khiển Word, Excel, PowerPoint bị từ chối, nên mọi tiêu chí "Chạy thật" và "Tự kiểm tra" dưới đây còn để ngỏ. Tác tử soát mã độc lập cũng bị dừng vì hết hạn mức model, nên mã mới chỉ qua kiểm thử máy và một lần chạy thử trong trình duyệt.

### Bảng nghiệm thu (tiêu chí ở `ke-hoach-v4.md` mục 5)

| Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|
| Bấm Dừng giữa chuỗi sửa, hỏi "anh vừa làm gì": AI kể đúng phần đã làm (Chạy thật trong Word) | **Chưa kiểm được** trong Office. Kiểm thử máy đạt: bước đã chạy còn kết quả, bước chưa chạy ghi "chưa thực hiện", có mốc dừng, lượt sau model nhận đủ; nút Hoàn tác vẫn có sau khi dừng | `test/unit/dot0b.test.ts` nhóm "lịch sử chỉ thêm" |
| Lượt thứ hai cùng hội thoại (key Anthropic): token đọc từ bộ nhớ đệm > 0, hiện trong khung (Chạy thật) | **Chưa kiểm được** với key thật. Kiểm thử máy đạt: yêu cầu có hai mốc `cache_control`, phần đầu hội thoại gửi lại giữ nguyên từng byte, lời nhắc hệ thống không đổi giữa hai lượt, đọc đúng `cache_read_input_tokens`. Dòng "Token: vào … · ra … · đọc từ bộ nhớ đệm …" hiện đúng trong trình duyệt với nhà cung cấp giả lập | `dot0b.test.ts`; `tasks/dot-0b-bang-chung/01` |
| Giả lập quá tải hai lần rồi thành công: tự thử lại, hiện "đang thử lại lần n", bấm Dừng hủy được (Kiểm thử máy) | **Đạt** — thử lại sau 2 và 6 giây, sự kiện báo "lần 1/3", "lần 2/3"; bấm Dừng trong lúc chờ thì hủy, không gọi thêm. Trong trình duyệt: hiện "Máy chủ nhà cung cấp lỗi 503: đang thử lại lần 1/3 sau 2 giây… Bấm Dừng để hủy." | `dot0b.test.ts` nhóm "tự thử lại"; `tasks/dot-0b-bang-chung/01` |
| Lệnh công cụ bị cắt dở: không được thực thi (Kiểm thử máy) | **Đạt** — cả ba kiểu API (`length`, `max_tokens`, `MAX_TOKENS`); model được báo chia nhỏ; cắt hai lần liên tiếp thì dừng lượt | `dot0b.test.ts` nhóm "lệnh công cụ bị cắt dở" |
| Nhận xét trong văn bản chứa câu ra lệnh: AI không làm theo và báo anh (Chạy thật) | **Chưa kiểm được**. Lời nhắc hệ thống đã có quy tắc kể rõ nguồn không tin cậy (Kiểm thử máy) | `dot0b.test.ts` nhóm "an toàn" |
| Ghi đè ô bảng Word đang có dữ liệu: có hỏi kèm lý do (Tự kiểm tra) | **Chưa kiểm được** — bước Tự kiểm tra đã viết nhưng chưa chạy trong Office. Kiểm thử máy cho phần tính ô bị ghi đè: đạt | `src/office/selftest.ts`; `dot0b.test.ts` |
| 7 mục chưa kiểm: tìm web; mở lại và xóa hội thoại; tự tóm tắt; bấm địa chỉ ô/đoạn/slide; nhập kỹ năng; đính kèm ảnh; lưu Chỉ dẫn riêng (Chạy thật, mỗi mục một ảnh) | **Chưa kiểm được** | — |
| Tự kiểm tra không giảm: Excel 22/22, Word 19/19 (17 cũ + 2 mới), PowerPoint 14/14 | **Chưa kiểm được** | — |
| `npm test` đạt, kiểm kiểu không lỗi, đóng gói đạt, phát hành tự động chạy | **Đạt** — 150/150, tsc sạch, gh-pages nhận bản 3.7.0 | Lệnh chạy trong phiên; `git log origin/gh-pages` |

### Việc đã làm theo 7 mục của kế hoạch

1. **Lịch sử chỉ thêm** — xong: `loop.ts` ghi từng tin nhắn ngay khi có, `Chat.tsx` lưu xuống máy theo từng bước; lời nhắc hệ thống cố định (`prompt.ts`), trạng thái tài liệu và kỹ năng nạp sẵn nằm trong tin nhắn của lượt; bỏ việc cắt ngắn kết quả cũ giữa các lượt. Chưa làm: giữ khối suy luận của model Claude gửi lại lượt sau (để Đợt 1B, cùng công tắc "Suy nghĩ kỹ").
2. **Tự thử lại và lệnh bị cắt dở** — xong; trần đầu ra Anthropic nâng từ 8.192 lên theo model, tối đa 32.000.
3. **Bộ nhớ đệm lời nhắc Anthropic, số token mỗi lượt** — xong về mã; chưa thấy số thật với key thật.
4. **Ngưỡng tóm tắt theo sức chứa model** — xong: 60% sức chứa, 2,5 ký tự mỗi token, sàn 8.000 và trần 400.000 ký tự; token thực đo vượt 75% sức chứa cũng tóm tắt.
5. **Tách điều phối khỏi `Chat.tsx`** — xong: `runTurn` trong `src/core/agent/turn.ts`, kiểm thử được ngoài Office.
6. **An toàn** — xong: quy tắc chỉ dẫn ẩn; ba cảnh báo rủi ro Word mới; hạn chờ 120 giây mỗi bước công cụ; key Gemini trong tiêu đề. Thư viện: 0 cảnh báo bảo mật ở bản chạy; không nâng các bản lớn (pdfjs 6, pptxgenjs 4, React 19) vì rủi ro hỏng biểu đồ PowerPoint và đọc PDF, không có lợi ích trước mắt.
7. **Kiểm nốt 7 mục chưa chạy thật** — chưa làm (cần điều khiển Office).

Việc còn lại từ Đợt 0A: mục 3 (báo đúng lý do khi model OpenRouter bị giới hạn nơi dùng, tránh chọn model đó làm mặc định) đã sửa, có kiểm thử máy; chưa đổi model mặc định của OpenAI vì chưa có key OpenAI chạy được để biết model nào không cần xác minh tổ chức. Mục 4: tệp sao lưu thử đã chuyển vào Thùng rác (anh duyệt). Lỗi nhỏ 5 (bảng Lịch sử che hội thoại khi gửi): đã sửa.

### Việc phiên sau phải làm trước khi sang Đợt 1A

Mở phiên trên máy, cho phép điều khiển Word, Excel, PowerPoint, rồi kiểm lần lượt các dòng "Chưa kiểm được" ở bảng trên; ảnh lưu vào `tasks/dot-0b-bang-chung/`. Thứ tự gợi ý: đầu khung hiện v3.7 → Tự kiểm tra ba ứng dụng → bốn bài Chạy thật trong Word bằng key Anthropic (Claude Haiku 5.5 cho rẻ) → 7 mục còn lại. Bài "tự tóm tắt" cần hội thoại dài hơn ngưỡng: dùng một nguồn có sức chứa nhỏ hoặc đọc một văn bản dài nhiều lượt.
