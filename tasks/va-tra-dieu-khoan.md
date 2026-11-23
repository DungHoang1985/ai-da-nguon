# Bản vá: tra "Điều N" khớp nhầm mục "N." của phụ lục (23/11/2026)

Nhánh `claude/va-tra-dieu-khoan`, tách từ `main@b2404be` (4.11.8). **Đã phát hành 4.11.9** (anh Dũng duyệt 23/11/2026): `main@81893e2` → `gh-pages@7188dc8`; Tự kiểm tra Word trên bản phát hành **41/41** (`tu-kiem-tra-word-phat-hanh-4119.json`).

## Lỗi

Phát hiện trong phép thử A/B (`tasks/ab-dinh-dang-gon.md`, lượt `khoi-H1-A-1.json`): trên bản 4.11.8, AI dừng lượt với câu "word_read_clause báo cùng một lỗi 2 bước liên tiếp", chờ 403 giây mà không trả lời.

Nguyên nhân: `findClause` ([src/core/outline.ts](../src/core/outline.ts)) chuẩn hóa "Điều 3" thành khóa "3", bỏ mất chữ "Điều", rồi so với mọi mục có khóa "3". Hợp đồng có Phụ lục 01 đánh số lại "1. … 2. … 3. Giờ hoạt động" nên "Điều 3" khớp cả Điều 3 (¶19–22) lẫn mục 3 của phụ lục (¶79–84). `word_read_clause` trả `ok: false` "có 2 điều khoản mang số". Model thường xoay sang `word_read_paragraphs`, tốn thêm một lần gọi AI. Khi model gọi lại y nguyên thì cơ chế chống lặp của vòng tác tử dừng lượt.

Mức ảnh hưởng: mọi hợp đồng có phụ lục đánh số lại từ 1. Trên bản 4.11.8, cả 12/12 lượt H1 và R1 trong bộ A/B (cả hai định dạng) đều phải đọc lại Điều 3 bằng "Đọc đoạn 19–23"; một lượt bị dừng hẳn.

## Sửa

Khi lời gọi đúng dạng "Điều N" / "Article N" (không kèm khoản, điểm) và văn bản có mục kiểu "Điều N" thì chỉ lấy các mục kiểu đó. Không đổi:
- Gọi số trần "3", hoặc "khoản 1 Điều 3", thì vẫn trả mọi mục trùng số để AI chọn.
- Phụ lục có điều mang chữ "Điều" riêng ("Điều 5. Giá phụ lục") thì vẫn báo trùng, như kiểm thử cũ.
- Văn bản không có mục "Điều N" nào (đánh số 1., 2. ngay từ đầu) thì "Điều 3" vẫn tìm ra mục "3.".

Cùng hàm này dùng cho tệp .docx đính kèm nên phần đó cũng được sửa theo.

## Kiểm

| Kiểm | Kết quả |
|---|---|
| Kiểm thử máy mới `dot1a.test.ts` "gọi 'Điều N' không khớp mục 'N.' của phụ lục" | Mã cũ: không đạt (trả ¶0 và ¶5). Mã mới: đạt |
| Toàn bộ kiểm thử máy, kiểm kiểu | 820/820, `tsc` sạch |
| Tra trên hợp đồng soát thật (`tim-dieu.ts`) | "Điều 3" → chỉ ¶19–22; "3", "3.1" vẫn báo trùng như cũ; hợp đồng 50 trang: "Điều 1", "Điều 2", "Điều 27", "Phụ lục 1" đúng |
| Chạy thật, desktop ẩn, Nemotron 3 Ultra free, 0 USD | Bảng dưới |

| Bài | 4.11.8 (6 lượt A/B mỗi bài) | Bản vá (3 lượt) |
|---|---|---|
| H1 hỏi 5 dữ kiện | 42/48 ở bản JSON; 3–4 lần gọi; trung vị 27.435 token vào; mọi lượt có "Đọc đoạn 19–23"; một lượt thử khói bị dừng | **48/48; 2 lần gọi mỗi lượt; 16.778–17.477 token vào (khoảng −37%)**; không lượt nào phải xoay |
| R1 rà soát Điều 3, 5, ghi nhận xét | 6/6 đạt; 3–6 nhận xét; mọi lượt phải "Đọc đoạn 19–2x" | 3/3 đạt; 5–6 nhận xét, tất cả đúng Điều 3/5, không sửa chữ; 2/3 lượt đọc thẳng Điều 3, 5 (lượt còn lại tự chọn đọc ¶19–31); số lần gọi 5–10, nhiều hơn vì ghi nhiều nhận xét hơn |

Bằng chứng: `tasks/va-tra-dieu-khoan-bang-chung/`. Một lượt H1 lỗi kết nối của bộ chạy ("fetch failed") đã chạy bù.

## Còn lại

- Mục cùng số giữa khoản phần chính và mục phụ lục ("3.1" so với "Phụ lục 01, mục 3.1") vẫn báo trùng như thiết kế cũ. Chưa thấy gây dừng lượt trong các lượt thử.
- Excel, PowerPoint không dùng hàm này nên không chạy lại Tự kiểm tra.

## Chạy thật bằng Claude trên bản phát hành 4.11.9 (anh Dũng cho chạy, 23/11/2026)

Chế độ Trả phí, nguồn OpenRouter (key có tín dụng của anh). Desktop ẩn, mỗi lượt một bản sao sạch. Chạy xong add-in đã được trả về chế độ Miễn phí. Bộ chạy `chay-claude.mjs`; bằng chứng ở `claude-haiku/`, `claude-sonnet/`.

| Model | Bài | Kết quả | Lần gọi | Token vào · ra | Các bước đọc |
|---|---|---|---|---|---|
| Claude Haiku 5.5 | H1 (2 lượt) | 16/16, 16/16 | 2, 2 | 34.856 · 2.169; 35.209 · 2.522 | Đọc thẳng Điều 3, 4, 7, 9, 11; không xoay sang đọc theo đoạn |
| Claude Haiku 5.5 | R1 (2 lượt) | Đạt, đạt: 6 nhận xét, tất cả đúng Điều 3/5, không sửa chữ | 5, 6 | 110.460 · 15.375; 140.382 · 15.217 | Đọc thẳng Điều 3, 5 (rồi đọc thêm các điều liên quan để đối chiếu) |
| Claude Sonnet 5.5 | H1 | 16/16 | 2 | 34.744 · 1.014 | Đọc thẳng Điều 3, 4, 7, 9, 11 |
| Claude Sonnet 5.5 | R1 | Đạt: 5 nhận xét, đúng Điều 3/5, không sửa chữ | 4 | 86.763 · 6.984 | — (phần đầu khung bị cắt khỏi bản ghi) |

Chi phí ước theo giá niêm yết của Anthropic, chưa trừ phần đọc từ bộ nhớ đệm (dòng token của khung không tách phần này với nguồn OpenRouter): Haiku 5.5 (0,10 / 0,50 USD mỗi triệu token vào / ra) khoảng 0,05 USD; Sonnet 5.5 (2 / 10 USD) khoảng 0,32 USD; **tổng khoảng 0,37 USD** (trần đặt trước 1 USD). Số trừ thật xem ở openrouter.ai → Activity.
