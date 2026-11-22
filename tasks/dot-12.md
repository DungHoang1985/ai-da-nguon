# Đợt 12 — kế hoạch (anh duyệt 22/11/2026)

Anh chọn ngày 22/11/2026 (sau khi phát hành 4.11.2) năm nhóm việc, theo thứ tự anh xếp. Bản này là kế hoạch: chưa viết mã. Mỗi phần xong thì kiểm, ghi bảng nghiệm thu ở cuối tệp; phát hành khi anh duyệt.

Hiện trạng đã đối chiếu mã (`main@c2a9693`):
- Tệp .docx và chữ dán vào được chép nguyên vào tin nhắn (`withAttachments`, `src/core/attachments.ts`), không đo trước.
- Vòng xoay Miễn phí đã bỏ nguồn khi `estimateTokens` vượt sức chứa model (`rotation.ts`), nhưng **hàm này chỉ đếm tin nhắn**: bỏ qua lời nhắc hệ thống và bộ công cụ (Word Miễn phí khoảng 18.600 ký tự, Excel 14.700, PowerPoint 14.000; Trả phí tới 30.000), không biết hạn mức token mỗi phút (Groq 8.000 token/phút → "quá lớn" sau khi đã gửi). Chế độ Trả phí không đo gì.
- Trần số bước một lượt: Excel 12, Word 12, PowerPoint 20 (Đợt 7).

## A. Đo token đầu vào trước khi gửi, chọn model đủ sức chứa (anh thêm)

Việc:
1. Một hàm đo đủ cả lượt: lời nhắc hệ thống + bộ công cụ đang gửi + lịch sử + tin nhắn mới + chữ tệp đính kèm (ước lượng ký tự / 3,2 như `do-token.ts`; số thật lấy từ `usage` khi có).
2. Hiện số đo trên khung trước khi gửi khi lớn (ví dụ từ 20.000 token): "Lượt này khoảng N token — vừa / không vừa model đang chọn (sức chứa M)".
3. Chế độ Miễn phí: vòng xoay dùng số đo đủ, bỏ trước các nguồn không chứa nổi (gồm hạn token mỗi phút đã biết của Groq) thay vì gửi rồi nhận lỗi; không nguồn nào chứa nổi thì nói rõ và gợi ý: đọc tệp theo đoạn (phần D), bớt tệp, hoặc chế độ Trả phí.
4. Chế độ Trả phí: model đang chọn không chứa nổi thì báo trước khi gửi, gợi ý model cùng nguồn có sức chứa lớn hơn (anh chọn: chỉ gợi ý hay tự đổi — câu hỏi 2).

Tiêu chí nghiệm thu:
- A1 Kiểm thử máy: số đo lệch không quá 15% so với `usage` thật của 3 lượt Chạy thật (Word có tệp .docx 50 trang, Excel, PowerPoint).
- A2 Miễn phí: hợp đồng thử 50 trang đính kèm → không còn lần gửi nào bị nguồn trả "quá lớn"; nguồn bị bỏ có lý do hiện trên khung. Chạy thật.
- A3 Trả phí: chọn model sức chứa nhỏ + tệp lớn → khung báo trước khi gửi, kèm gợi ý. Chạy thật hoặc Tự kiểm tra.
- A4 Không đổi lời nhắc hệ thống, không đổi lịch sử đã gửi (bộ nhớ đệm lời nhắc giữ nguyên — kiểm thử chặn `dot0b` vẫn xanh).

## B. Excel còn trừ điểm trong buổi so sánh 22/11

1. **X3** — AI báo đã định dạng ngày dd/mm/yyyy, đọc lại thấy chưa. Thăm dò trước: ô ngày là chữ hay số; `numberFormat` ghi được chưa. Sửa công cụ hoặc kết quả trả về cho đúng sự thật (ví dụ báo "ô là chữ, định dạng ngày không áp được — cần đổi thành ngày trước").
2. **X6** — chạm trần 12 bước nên mất khoảng 15 phút và phải bấm Làm tiếp. Xem nhật ký bước: gộp lệnh định dạng (một lệnh nhiều vùng) hay nâng trần Excel (câu hỏi 3).
3. **X5** — kiểm toán mô hình nêu đủ 5 lỗi nhưng báo nhầm 2 chỗ: tìm mẫu báo nhầm trong `modelaudit.ts`, thêm kiểm thử ca phải bắt và ca dễ báo nhầm (quy tắc của `modelaudit.ts`).

Tiêu chí: B1 Tự kiểm tra Excel đạt đủ (hiện 40/40) có bước mới cho X3; B2 Chạy thật lại X3, X5, X6 bằng OpenRouter Nemotron 3 Ultra (Miễn phí, Nghĩ kỹ) bằng bộ chạy so sánh: X3 định dạng ngày đúng khi đọc lại, X5 không báo nhầm, X6 không phải Làm tiếp hoặc dưới 8 phút; chấm lại theo `dap-an.json`.

## C. Dọn HANDOFF và soát mã độc lập

1. Viết lại mục "Bước tiếp theo" (đang ghi đề xuất ngày 18/11, nhiều việc đã xong) và "Vấn đề mở" (còn ghi chữ “ư” chưa đo, 4.7.2 chưa phát hành…) cho đúng hiện trạng.
2. Soát mã độc lập (một tác tử không thấy kết luận của tôi) cho: mã ngày 22/11 (bảng PowerPoint, cảnh báo Làm tiếp), mã Đợt 7 (chưa từng soát), và mã của phần A, B, D khi xong. Sửa hết mục cao và trung bình; mục để lại ghi lý do.

Tiêu chí: C1 HANDOFF không còn việc đã xong nằm ở Vấn đề mở / Bước tiếp theo; C2 có báo cáo `soat-ma-dot-12.md`, mọi mục cao / trung bình đã sửa và chạy lại kiểm thử.

## D. Tệp .docx đính kèm đọc theo đoạn

Làm giống PDF ở Đợt 5: tệp .docx lớn không chép nguyên vào tin nhắn mà gửi mục lục điều khoản (`outline.ts`) + số đoạn; AI đọc phần cần bằng một công cụ chỉ đọc (theo điều / theo dải đoạn). Tệp nhỏ (dưới ngưỡng, ví dụ 6.000 token) vẫn gửi nguyên như cũ. So sánh phiên bản (`word_compare_versions`) không đổi.

Tiêu chí: D1 Kiểm thử máy cho ngưỡng, mục lục, đọc theo điều; D2 Tự kiểm tra Word có bước đọc tệp đính kèm theo điều; D3 Chạy thật: hợp đồng thử 50 trang đính kèm, hỏi "Điều 6.4 quy định gì" bằng Nemotron Miễn phí — trả lời đúng, token đầu vào giảm so với gửi nguyên (ghi số trước / sau).

## E. Tính năng mới mục D: bộ nhớ dài hạn giữa các hội thoại

Phần này chỉ viết thiết kế (`tasks/dot-12-bo-nho.md`): nhớ gì (sở thích trình bày, thuật ngữ, tên dự án, quy ước của anh), lưu ở đâu (kho cục bộ của add-in, không lên mạng), AI đọc / ghi thế nào, cách anh xem và xóa, ảnh hưởng tới bộ nhớ đệm lời nhắc. **Anh duyệt thiết kế rồi mới viết mã.** Connectors (MCP), Outlook, hàm `=AI()` trong Excel để sau.

## Thứ tự và cách làm

Làm theo thứ tự anh xếp: A → B → C → D → E. Riêng D có liên quan A (D giảm token, A đo token): làm A trước để có số đo trước / sau cho D. Mỗi phần: thăm dò trong Office trên desktop ẩn nếu chạm Office.js, kiểm thử máy, Tự kiểm tra trên bản cục bộ (`cdp.mjs serve`), Chạy thật với OpenRouter Nemotron (Miễn phí, 50 lượt mỗi ngày). Không phát hành khi anh chưa duyệt.

## Anh đã quyết (22/11/2026)

1. Duyệt kế hoạch, làm theo thứ tự A → B → C → D → E.
2. Phần A, Trả phí: chỉ báo và gợi ý model lớn hơn, không tự đổi.
3. Phần B, X6: gộp lệnh cho ít bước hơn, không nâng trần.
4. Phát hành: mỗi phần xong là một bản vá, phát hành khi anh duyệt.

## Nghiệm thu

### Phần A — đo token đầu vào (làm 22/11/2026, nhánh `claude/sua-bang-ppt-va-canh-bao`, chưa phát hành)

Đã làm: `src/core/inputsize.ts` (đo đủ lời nhắc + công cụ + lịch sử + tệp; hai hệ số chữ / JSON; hệ số riêng cho họ Claude; gợi ý model đủ sức chứa; câu báo); vòng xoay Miễn phí đo theo họ model của từng nguồn và tính cả bộ công cụ; `runTurn` đo lần gửi đầu, phát sự kiện `input_size`, chế độ Trả phí mà model không chứa nổi thì KHÔNG gửi, báo lỗi `too_large` kèm gợi ý (anh chọn "chỉ báo và gợi ý"); khung hiện dòng số đo khi lượt từ 20.000 token, nút "Đổi model" ở lỗi quá lớn; mọi nguồn miễn phí đều quá nhỏ thì câu lỗi nói cách làm nhỏ lượt. Đã bỏ khỏi kế hoạch: bảng hạn token mỗi phút của Groq (con số theo từng model, chưa đối chiếu được tài liệu chính thức; vòng xoay vẫn bỏ nguồn sau lần bị từ chối "quá lớn" như trước).

| Tiêu chí | Kết quả | Bằng chứng (`tasks/dot-12-bang-chung/`) |
|---|---|---|
| A1 Số đo lệch ≤ 15% so với `usage` thật (Nemotron, 3 lượt) | Đạt | Một hệ số (3) lệch −14,9% · −11,4% · +21,2% (`PA1-D1`, `XA1-D1`, `WA1-D1`) → hai hệ số: PowerPoint 6.420 / 6.952 (−7,7%), Excel 5.763 / 5.686 (+1,4%), Word + hợp đồng 50 trang 41.653 / 39.039 (+6,7%) (`PA1-D2`, `XA1-D2`, `WA1-D2`) |
| A1 bổ sung: họ Claude | Đạt | Hệ số chung đếm thiếu 29–54% với Haiku 5.5 (`PA1-C3`, `PA2-C3`, `PC3-C1`); hệ số riêng: 11.490 / 11.629 (−1,2%), 24.442 / 23.378 (+4,6%) (`PA1-C4`, `PA2-C4`). Gemini chưa đo riêng được (lượt chọn Gemini ra đúng số của Nemotron — nghi vòng xoay dùng OpenRouter trước), dùng hệ số chung |
| A2 Miễn phí, hợp đồng 50 trang đính kèm: không lần gửi nào bị "quá lớn" | Đạt | `WA1-D2`: Nemotron trả lời đúng Điều 18.2(b), không đổi nguồn, không lỗi; khung hiện "Lượt này khoảng 41.653 token đầu vào" |
| A3 Trả phí, model không chứa nổi: báo trước khi gửi, kèm gợi ý | Đạt | `PC3-C2`: OpenRouter Llama 3.1 8B (131.072) + 887.000 ký tự → "Chưa gửi: lượt này khoảng 254.640 token … Model cùng nguồn chứa được: Nova Lite, Nova Pro, Llama 4 Scout", nút "Đổi model", không có dòng token (không gửi) |
| A4 Không đổi lời nhắc hệ thống, không đổi lịch sử | Đạt | Kiểm thử máy 723/723 gồm `dot0b` (bộ nhớ đệm) và `test/unit/dot12.test.ts` (lịch sử giữ nguyên khi chặn) |

Phát hành **4.11.3** ngày 22/11/2026 (anh duyệt): `main@aecac6c` → `gh-pages@3934dc6`. Trên bản phát hành: Chạy thật PowerPoint bằng Nemotron ước 6.420 / thật 6.952 token (`PA1-D5-phat-hanh-4113.json`); Tự kiểm tra PowerPoint 26/26 (`tu-kiem-tra-ppt-phat-hanh-4113.json`).

Sự cố trong lúc kiểm: lần thử A3 đầu tiên chọn Claude Haiku 5.5, tưởng sức chứa nhỏ nhưng thật là 1.000.000 token, nên lượt 887.000 ký tự được gửi thật: 554.610 token vào key Anthropic của anh (`PC3-C1`, dưới 1 USD theo giá Haiku). Lần sau chọn model sức chứa nhỏ đã biết trước khi gửi.
