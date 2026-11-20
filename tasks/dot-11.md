# Đợt 11 – Rút gọn token mỗi lượt Excel và PowerPoint

Anh bảo "làm tiếp Đợt 11" ngày 20/11/2026, theo đề xuất ở cuối Đợt 10: làm Excel rồi PowerPoint theo đúng cách đã làm cho Word. Nhánh `claude/dot-11` (worktree `dot-11-tiep-tuc-0edf71`), tách từ `main@7cf5786` (bản 4.10.0).

## 1. Cách làm (giữ nguyên lựa chọn anh đã chốt ở Đợt 10, mục 6 của `tasks/dot-10.md`)

- **A. Viết gọn cho mọi chế độ:** rút mô tả công cụ và mô tả tham số, gộp các dòng chỉ đường tới công cụ chuyên trong lời nhắc thành một dòng, bỏ phần chuẩn làm việc trùng với lời nhắc, thêm `brief` (≤ 200 ký tự) cho kỹ năng Excel, PowerPoint trong danh mục kỹ năng. Không rút quy tắc an toàn, không đổi hành vi công cụ.
- **B. Nạp công cụ theo việc, chỉ ở chế độ Miễn phí:** nhóm lõi luôn gửi; nhóm chuyên bật theo từ khóa, kỹ năng, tệp đính kèm, `load_tools`, hoặc khi model đã dùng nhóm đó trong hội thoại. Trả phí vẫn gửi đủ công cụ.
- Model thử: Claude Haiku 5.5 cho bài trước / sau (Trả phí); một nguồn miễn phí cho bài Miễn phí.

## 2. Số đo mốc (bản 4.10.0, `kiem-thu/do-token.ts`)

| Ứng dụng | Công cụ | Lời nhắc (ký tự) | Công cụ (ký tự) | Tổng |
|---|---|---|---|---|
| Excel | 33 | 8.070 (riêng Excel 2.456 · chuẩn làm việc 1.719 · danh mục kỹ năng 1.960 · quy tắc chung 1.935) | 23.684 | **31.754** |
| PowerPoint | 27 | 7.498 (riêng PowerPoint 2.705 · chuẩn làm việc 1.365 · danh mục kỹ năng 1.493 · quy tắc chung 1.935) | 15.795 | **23.293** |

(Số ở HANDOFF — Excel 32.758, PowerPoint 24.102 — là của bản 4.9.0; Đợt 10 đã rút phần dùng chung.) Tệp: `dot-11-bang-chung/do-token-*-moc-4100.json`.

Khác Word: mô tả công cụ Excel, PowerPoint vốn đã ngắn; phần lớn là lược đồ tham số. Viết gọn (A) vì thế ước giảm ít hơn Word; phần giảm chính sẽ đến từ lời nhắc (dòng chỉ đường lặp tên công cụ, chuẩn làm việc trùng lời nhắc, danh mục kỹ năng chưa có `brief`).

## 3. Nhóm công cụ (B, đề xuất)

**Excel** — lõi: tổng quan, đọc, ghi, định dạng, thêm sheet, tạo bảng, sắp xếp, xóa vùng, chọn vùng, tìm, quét lỗi, lần công thức, nhận xét ô, chèn/xóa hàng cột, thao tác sheet, xem đối tượng có sẵn, kỹ năng, đọc trang tệp đính kèm.

| Nhóm | Công cụ |
|---|---|
| `bieu_do` biểu đồ | `excel_create_chart`, `excel_edit_chart` |
| `pivot` PivotTable | `excel_create_pivot`, `excel_edit_pivot` |
| `loc_dinh_dang` lọc, định dạng có điều kiện, danh sách chọn | `excel_filter`, `excel_conditional_format`, `excel_data_validation` |
| `in_an` thiết lập in | `excel_page_setup` |
| `mo_hinh` mô hình tài chính | `excel_audit_model`, `excel_named_range`, `excel_what_if` |
| `lam_sach_khoa` làm sạch dữ liệu, khóa ô | `excel_clean_data`, `excel_protect` |
| `ngoai` ứng dụng khác | `read_other_app`, `write_other_app` |

**PowerPoint** — lõi: tổng quan, thêm slide (một và nhiều), sửa chữ, ô chữ, hình, bảng, chuyển slide, sửa / xóa hình, định dạng chữ, đọc vùng chọn, chụp slide, đọc mẫu, kỹ năng, đọc trang tệp đính kèm.

| Nhóm | Công cụ |
|---|---|
| `bieu_do_so_do` biểu đồ, sơ đồ quy trình | `ppt_add_chart`, `ppt_add_process_flow` |
| `sap_xep` sắp xếp slide, đổi bố cục | `ppt_move_slide`, `ppt_duplicate_slide`, `ppt_delete_slide`, `ppt_set_layout` |
| `kiem_bai` kiểm cả bài | `ppt_check_deck` |
| `hinh_anh` chèn ảnh đính kèm | `ppt_insert_image` |
| `bang` sửa bảng có sẵn | `ppt_edit_table` |
| `ngoai` ứng dụng khác | `read_other_app`, `write_other_app` |

Ảnh đính kèm bật `hinh_anh` (PowerPoint); tệp .docx đính kèm chỉ bật nhóm `soat` ở Word như cũ.

## 4. Tiêu chí nghiệm thu

| # | Tiêu chí | Cách kiểm |
|---|---|---|
| T1 | Phần cố định mỗi lượt: Trả phí giảm ≥ 20% (Excel, PowerPoint); Miễn phí nhóm lõi giảm ≥ 45%, so với mốc 4.10.0 | `kiem-thu/do-token.ts`, JSON trước / sau; kiểm thử chặn phình lại |
| T2 | Token thật mỗi lần gọi Claude Haiku 5.5 (Trả phí, một bài ngắn mỗi ứng dụng) giảm ≥ 15% | Chạy thật, dòng Token trên khung |
| T3 | Bộ bài Excel và PowerPoint bằng Haiku (Trả phí) làm xong đúng, số lần gọi không tăng quá 1 so với 4.10.0; mỗi ứng dụng một bài Miễn phí làm xong nhờ nhóm tự bật | Chạy thật trên desktop ẩn |
| T4 | Miễn phí: mỗi ứng dụng một bài cố ý không khớp từ khóa → model gọi `load_tools` rồi làm xong | Chạy thật |
| T5 | Kiểm thử máy, kiểm kiểu sạch; Tự kiểm tra Word 39/39, Excel 40/40, PowerPoint 25/25 | `npm test`, `tsc`, nút Tự kiểm tra trên desktop ẩn |
| T6 | Soát mã độc lập, sửa hết mục cao và trung bình | `soat-ma-dot-11.md` |

Ngưỡng T1, T2 thấp hơn Word (25% / 50% / 20%) vì mô tả công cụ Excel, PowerPoint vốn đã ngắn (mục 2); nếu đo được cao hơn thì ghi số thật.

## 5. Đã làm

(cập nhật dần)
