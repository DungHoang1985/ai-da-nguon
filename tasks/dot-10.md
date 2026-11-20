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

## 7. Trạng thái

- 20/11: đo mốc xong (mục 2); anh chốt mục 6. Đang làm: bước A (viết gọn).
