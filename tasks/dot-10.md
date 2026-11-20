# Đợt 10 – Rút gọn token mỗi lượt (Word trước)

Anh chọn ngày 20/11/2026 (phiên "tiếp tục dự án"). Nhánh `claude/dot-10`, tách từ `main@b23ddd8` (bản 4.9.0).

## 1. Vì sao làm

- Mỗi lần gọi AI trong Word gửi lại toàn bộ lời nhắc hệ thống, 35 công cụ và trạng thái tài liệu. Hậu quả:
  - Groq miễn phí (8.000 token mỗi phút) lỗi 413 trong Word. Chế độ Miễn phí phải bỏ qua Groq; Đợt 9 Chạy thật đi theo vòng Gemini → Groq 413 → OpenRouter.
  - Nguồn không có bộ nhớ đệm (OpenRouter `:free`, Groq, Cohere, phần lớn nguồn tự khai) tốn trọn số token đó ở **mỗi** lần gọi; nguồn miễn phí hết lượt nhanh hơn.
  - Lượt dài phải tóm tắt hội thoại sớm hơn.

## 2. Số đo mốc (bản 4.9.0)

Bài đo: `kiem-thu/do-token.ts` (dựng đúng lời nhắc hệ thống và bộ công cụ như khung add-in, danh mục kỹ năng lọc theo ứng dụng; đếm ký tự JSON theo dạng gửi Anthropic).

| Ứng dụng | Công cụ | Lời nhắc (ký tự) | Công cụ (ký tự) | Tổng (ký tự) |
|---|---|---|---|---|
| Word | 35 | 11.068 (riêng Word 4.306 · chuẩn làm việc 1.247 · danh mục kỹ năng 3.580 · quy tắc chung 1.935) | 29.854 | 40.922 |
| Excel | 33 | 8.484 | 24.274 | 32.758 |
| PowerPoint | 27 | 7.717 | 16.385 | 24.102 |

- Số token thật [Đã đo, Đợt 9, `tasks/dot-9-bang-chung/chay-that-tra-phi.json`]: Claude Haiku 5.5, văn bản ngắn, 2 lần gọi = 37.207 token vào, tức khoảng **18.600 token mỗi lần gọi**. Chữ Việt có dấu tốn token hơn chữ Anh (khoảng 2,2 ký tự mỗi token với bộ đếm của Claude).
- Công cụ Word nặng nhất: `word_create_admin_document` 3.514 · `word_format_text` 1.591 · `word_replace_text` 1.427 · `word_insert_paragraphs` 1.314 · `word_insert_clause` 1.290 · `word_read_paragraphs` 1.184 · `write_other_app` 1.181 · `word_fill_form` 1.129.

## 3. Hai cách giảm (đề xuất)

**A. Viết gọn** (an toàn, không đổi hành vi): rút mô tả công cụ và lược đồ tham số (bỏ câu lặp lại giữa các công cụ, chuyển dặn dò chung lên lời nhắc một lần, mô tả tham số ngắn), gọn lời nhắc riêng Word và danh mục kỹ năng (mô tả kỹ năng ≤ 150 ký tự). Ước giảm 25–35%.

**B. Nạp công cụ theo việc**: mỗi lượt chỉ gửi nhóm **lõi** (tổng quan, mục lục, tìm, đọc, sửa, chèn, xóa, định dạng, nhận xét, chuyển tới, kỹ năng). Các nhóm **chuyên** chỉ gửi khi cần:

| Nhóm chuyên | Công cụ | Ký tự hiện tại |
|---|---|---|
| Văn bản hành chính | `word_create_admin_document`, `word_set_page_format`, `word_check_page_format` | ~5.200 |
| Điền mẫu | `word_list_form_fields`, `word_fill_form` | ~1.900 |
| Soát và so sánh | `word_check_consistency`, `word_compare_versions`, `word_read_parts` | ~3.100 |
| Nhận xét và sửa đổi theo dõi | `word_list_comments`, `word_reply_comment`, `word_resolve_comment`, `word_list_revisions`, `word_review_revision` | ~2.200 |
| Bảng | `word_read_tables`, `word_insert_table`, `word_fill_table` | ~1.800 |
| Ứng dụng khác, tệp đính kèm | `read_other_app`, `write_other_app`, `read_attachment_pages` | ~2.500 |

- Bật nhóm: (1) theo từ khóa tiếng Việt không dấu trong yêu cầu, dùng lại cách tự chọn kỹ năng đã có; (2) kỹ năng đã nạp thì kéo theo nhóm của nó (vd `dien-mau` → Điền mẫu); (3) có tệp đính kèm thì bật nhóm tệp; (4) lời nhắc có một dòng liệt kê các nhóm chưa nạp và một công cụ nhỏ `load_tools(group)` để model tự xin khi cần.
- Nhóm đã bật thì **giữ tới hết hội thoại** (không tắt lại), để bộ công cụ chỉ đổi theo chiều thêm: nguồn có bộ nhớ đệm (Anthropic) chỉ mất bộ nhớ đệm đúng một lần mỗi lần thêm nhóm.
- Ước giảm cùng với A: phần cố định của Word từ ~41.000 xuống ~17.000–20.000 ký tự (giảm 50–60%).
- Rủi ro: model không biết công cụ chuyên có sẵn → làm đường vòng. Giảm bằng dòng liệt kê nhóm + `load_tools` + bộ bài Chạy thật ở mục 5.

**Không hứa trước:** Groq miễn phí 8.000 token mỗi phút. Sau A + B một lượt Word với văn bản ngắn ước còn khoảng 8.000–9.000 token, nên Groq **có thể vẫn chưa chạy được**; sẽ đo và ghi đúng kết quả.

## 4. Bất biến không làm trái

- Lời nhắc hệ thống không chứa thứ hay đổi; bộ công cụ không đổi giữa các lần gọi **trong một lượt** (chỉ thêm nhóm ở đầu lượt hoặc sau `load_tools`) — kiểm thử chặn.
- Quy tắc an toàn (`SAFETY_RULES`), thao tác rủi ro luôn hỏi, Track Changes: giữ nguyên chữ, không rút.
- Gemini chỉ nhận enum là chuỗi: `load_tools` dùng enum chuỗi; Chạy thật ít nhất một model Gemini.
- Tự kiểm tra gọi thẳng `execute` nên không phụ thuộc nhóm nạp; mọi công cụ vẫn chạy được khi model gọi (kể cả công cụ chưa nạp, nếu model đoán đúng tên).

## 5. Tiêu chí nghiệm thu (đề xuất)

| # | Tiêu chí | Cách kiểm |
|---|---|---|
| T1 | Phần cố định của lượt Word (lời nhắc + công cụ, nhóm lõi) giảm ≥ 50% so với mốc 40.922 ký tự | `kiem-thu/do-token.ts`, ghi JSON trước / sau |
| T2 | Token thật mỗi lần gọi (Claude Haiku 5.5, cùng bài chèn đoạn Đợt 9) giảm ≥ 40% so với ~18.600 | Chạy thật, đọc dòng Token trên khung |
| T3 | Không giảm chất lượng: 6 bài Chạy thật trong Word (rà soát hợp đồng ghi nhận xét · xử lý Track Changes · điền mẫu · soạn công văn · soát nhất quán · so phiên bản) đều làm xong đúng; số lần gọi AI mỗi bài không tăng quá 1 so với 4.9.0 | Chạy thật trên desktop ẩn, so bảng trước / sau cùng model |
| T4 | Model tự xin nhóm chưa nạp: một bài cố ý không khớp từ khóa mà cần công cụ chuyên → model gọi `load_tools` rồi làm xong | Chạy thật |
| T5 | Chế độ Miễn phí: ít nhất một nguồn miễn phí làm xong bài Word ngắn; ghi Groq có hết 413 không (Đạt nếu hết, nếu không ghi số token còn thiếu) | Chạy thật |
| T6 | Kiểm thử máy và kiểm kiểu sạch; Tự kiểm tra Word 39/39, Excel 40/40, PowerPoint 25/25 | `npm test`, `tsc`, nút Tự kiểm tra trên desktop ẩn |
| T7 | Soát mã độc lập, sửa hết mục cao và trung bình | Báo cáo `soat-ma-dot-10.md` |

## 6. Anh chốt (20/11/2026)

1. **Cách làm:** viết gọn (A) cho mọi chế độ; **nạp theo việc (B) chỉ bật ở chế độ Miễn phí**. Chế độ Trả phí vẫn gửi đủ công cụ (Claude có bộ nhớ đệm nên phần cố định rẻ, và không có rủi ro AI không thấy công cụ chuyên).
2. **Phạm vi:** Word trước, kiểm và phát hành 4.10.0; Excel, PowerPoint làm đợt sau theo cùng cách.
3. **Model thử:** Claude Haiku 5.5 cho bộ bài trước / sau; thêm một nguồn miễn phí (Gemini hoặc OpenRouter) cho bài chế độ Miễn phí và bài Groq.

Tiêu chí điều chỉnh theo lựa chọn trên (thay mục 5 ở các điểm sau):

| # | Tiêu chí |
|---|---|
| T1 | Phần cố định của lượt Word: chế độ Trả phí (A) giảm ≥ 25%; chế độ Miễn phí, nhóm lõi (A + B) giảm ≥ 50% so với 40.922 ký tự |
| T2 | Token thật mỗi lần gọi với Claude Haiku 5.5 (Trả phí, bài chèn đoạn Đợt 9) giảm ≥ 20% so với ~18.600 |
| T3 | 6 bài Word Chạy thật bằng Haiku (Trả phí) làm xong đúng, số lần gọi không tăng quá 1 so với 4.9.0; thêm 2 bài ở chế độ Miễn phí (điền mẫu, soạn công văn) bằng nguồn miễn phí làm xong nhờ nhóm tự bật |
| T4 | Ở chế độ Miễn phí, bài cố ý không khớp từ khóa → model gọi `load_tools` rồi làm xong |

## 7. Đã làm

- **A (mọi chế độ):** bỏ câu giải thích `expect` lặp ở 10 công cụ (lời nhắc đã giải thích một lần); rút mô tả và tham số của 35 công cụ; 5 dòng hướng dẫn công cụ chuyên trong lời nhắc Word gộp thành một dòng `WORD_SPECIAL_LINE`, phần hành vi riêng ("báo đủ mọi điểm theo total", "tóm tắt đủ theo differences") chuyển vào mô tả của chính công cụ đó; chuẩn làm việc Word bỏ phần trùng với lời nhắc Word; 11 kỹ năng Word có `brief` (≤ 200 ký tự) cho danh mục kỹ năng. Excel, PowerPoint chỉ đổi theo phần dùng chung (`read_other_app`, `write_other_app`, `read_attachment_pages`, câu đầu danh mục kỹ năng, `brief` của 3 kỹ năng dùng chung).
- **B (chỉ Miễn phí, chỉ Word):** `src/core/toolgroups.ts` — 6 nhóm chuyên (`hanh_chinh`, `dien_mau`, `soat`, `nhan_xet_sua_doi`, `bang`, `ngoai`), nhóm lõi 16 công cụ. Nhóm bật khi: lời người dùng khớp từ khóa (chuỗi có dấu so trên câu gốc để "bảng" không lẫn "bằng"), kỹ năng tự chọn kéo theo, tệp PDF / .docx đính kèm, model gọi `load_tools`, model gọi thẳng công cụ của nhóm hoặc `use_skill` một kỹ năng của nhóm. Nhóm đang bật được suy ra lại từ lịch sử ở mỗi bước (không lưu trạng thái riêng), nên các lượt sau của cùng hội thoại vẫn có nhóm đã dùng. Lời nhắc thêm mục "Nhóm công cụ chuyên" (cố định trong hội thoại).
- Bài đo `kiem-thu/do-token.ts` (`--free` cho nhóm lõi); bài thử `kiem-thu/cdp-bai/dot-10-moi.js`; 16 kiểm thử ở `test/unit/dot10.test.ts`, gồm kiểm thử chặn phình lại (Trả phí ≤ 30.700, Miễn phí ≤ 20.400 ký tự).

## 8. Kết quả (bản cục bộ 4.10.0, desktop ẩn, 20/11/2026)

### Số token thật, Claude Haiku 5.5, chế độ Trả phí (cùng bài, cùng tài liệu, hội thoại mới mỗi bài)

"Token vào" là tổng đầu vào mọi lần gọi của bài (gồm phần đọc từ và ghi vào bộ nhớ đệm). Tệp: `truoc-*.json` (bản 4.9.0 đã phát hành), `sau-*.json` (bản cục bộ 4.10).

| Bài | 4.9.0: lần gọi · token vào | 4.10: lần gọi · token vào | Giảm | Kết quả |
|---|---|---|---|---|
| B1 Soát nhất quán trước khi ký | 2 · 50.100 | 2 · 41.587 | 17% | Cả hai bản đều tìm đủ 9 lỗi cài sẵn |
| B2 So với bản v2 đính kèm | 2 · 49.441 | 2 · 40.937 | 17% | Cả hai kê đủ các chỗ khác theo điều |
| B3 Đối tác sửa gì, nhận xét gì | 3 · 61.940 | 3 · 49.145 | 21% | Cùng 4 bước đọc, cùng kết luận |
| B4 Rà soát Điều 3, 5, ghi nhận xét | 6 · 142.104 | 4 · 69.023 | 51% | 4.9.0: 5 nhận xét (1 lần lỗi `expect` phải ghi lại); 4.10: 3 nhận xét gộp có mức Vàng/Đỏ, cùng các rủi ro chính |
| B5 Chèn một đoạn cuối văn bản | 2 · 38.984 | 2 · 30.475 | **22%** | Đúng |
| B6 Điền hợp đồng đặt cọc | 3 · 65.759 | 3 · 52.422 | 20% | 15/15 chỗ cả hai bản |
| B7 Soạn công văn | 2 · 47.132 | 2 · 39.050 | 17% | Một lệnh dựng đúng thể thức cả hai bản |

Mỗi lần gọi (B5): 19.492 → 15.238 token.

### Chế độ Miễn phí (nạp theo việc)

| Bài | Nguồn | Lần gọi · token vào | Kết quả |
|---|---|---|---|
| T4: "Dòng chữ lặp lại ở trên cùng và dưới cùng mỗi trang ghi gì?" (không khớp từ khóa nào) | Gemini 3.8 Flash | 3 · 19.707 | Model gọi `load_tools(soat)` → `word_read_parts` → trả đúng đầu trang, chân trang |
| Điền hợp đồng đặt cọc | Gemini 3.8 Flash | 3 · 24.684 | 15/15 chỗ (nhóm `dien_mau` bật theo từ khóa) |
| Soạn công văn | Gemini im lặng → Groq "yêu cầu quá lớn" → OpenRouter Nemotron 3 Ultra free | 3 · 32.211 | Dựng đúng thể thức, còn 2 chỗ cần điền |
| Chèn một đoạn | Groq GPT OSS 120B → Gemini | 2 · 10.240 | Lần gọi đầu của **Groq chạy được** (trước đây lỗi 413 ngay lần đầu) và ghi đúng; lần gọi thứ hai chạm hạn 8.000 token/phút của Groq → tự đổi sang Gemini làm xong |

### Bảng nghiệm thu

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| T1 | Phần cố định Word: Trả phí giảm ≥ 25%, Miễn phí nhóm lõi giảm ≥ 50% | **Đạt** — Trả phí 40.922 → 30.049 ký tự (−26,6%); Miễn phí 18.135 (−55,7%) | `do-token-word-moc-490.json`, `do-token-word-tra-phi.json`, `do-token-word-mien-phi.json`; kiểm thử chặn phình lại |
| T2 | Token thật mỗi lần gọi (Haiku, bài chèn đoạn) giảm ≥ 20% | **Đạt** — 19.492 → 15.238 (−21,8%) | `truoc-B5-chen.json`, `sau-B5-chen.json` |
| T3 | 7 bài Haiku làm xong đúng, số lần gọi không tăng quá 1; 2 bài Miễn phí làm xong | **Đạt** — 7/7 xong, số lần gọi bằng hoặc ít hơn; điền mẫu và công văn ở Miễn phí làm xong | bảng trên, `sau-*.json`, `mien-phi-B6-*.json`, `mien-phi-B7-*.json` |
| T4 | Miễn phí: bài không khớp từ khóa → model gọi `load_tools` rồi làm xong | **Đạt** (Gemini 3.8 Flash) | `mien-phi-T4-load-tools.json` |
| T5 | Ít nhất một nguồn miễn phí làm xong bài Word ngắn; ghi Groq có hết 413 không | **Đạt** — Groq hết 413 ở bài ngắn (lần gọi đầu chạy được); vẫn bị hạn 8.000 token/phút ở lần gọi thứ hai và vẫn "quá lớn" khi đã nạp nhóm văn bản hành chính → vòng xoay tự đổi nguồn | `mien-phi-T5-groq.json`, `mien-phi-B7-cong-van.json` |
| T6 | Kiểm thử máy, kiểm kiểu; Tự kiểm tra Word 39/39, Excel 40/40, PowerPoint 25/25 | **Đạt** — 663/663, tsc sạch; Tự kiểm tra Word 39/39, Excel 40/40, PowerPoint 25/25 trên bản 4.10 cục bộ | `tu-kiem-tra-*.json` |
| T7 | Soát mã độc lập, sửa hết mục cao và trung bình | Đang chạy | `soat-ma-dot-10.md` |

Ghi chú: một lần chạy "sau" của B7 bị bỏ vì quên mở lại Word (văn bản còn công văn của lần chạy mốc); đã chạy lại trên văn bản trống. Câu trả lời của AI ở lần chạy hỏng có nhắc tên tài khoản Office của anh; tệp đó đã bị bản chạy lại ghi đè, đã `grep` lại toàn bộ bằng chứng: không còn tên.

## 9. Trạng thái

- 20/11: A + B xong, kiểm xong T1–T6. Đang chạy soát mã độc lập (T7), sau đó phát hành 4.10.0.
