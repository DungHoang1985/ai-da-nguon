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

- **A (mọi chế độ):** rút mô tả và tham số của 33 công cụ Excel, 27 công cụ PowerPoint (mô tả tham số dùng chung `addressProp`, `sheetProp`, `slideProp`, `shapeProp` lặp ở hầu hết công cụ); lời nhắc Excel, PowerPoint viết lại thành `EXCEL_SYSTEM_PROMPT`, `PPT_SYSTEM_PROMPT` với một dòng chỉ đường `EXCEL_SPECIAL_LINE`, `PPT_SPECIAL_LINE` (đặt ở `src/core/toolgroups.ts`, như `WORD_SPECIAL_LINE`); chuẩn làm việc Excel, PowerPoint bỏ phần trùng lời nhắc; `brief` cho 8 kỹ năng (5 Excel/PowerPoint riêng, `thiet-ke-bieu-do`, `kiem-tra-bai-trinh-chieu`, `trinh-chieu-tu-tai-lieu`) — nay mọi kỹ năng dựng sẵn đều có `brief`.
- **B (chỉ Miễn phí):** `EXCEL_TOOL_GROUPS` (7 nhóm), `PPT_TOOL_GROUPS` (6 nhóm) như mục 3; `toolGroupsFor` trả nhóm của cả ba ứng dụng ở chế độ Miễn phí; `freeBasePrompt(base, groups)` thay đúng dòng chỉ đường của ứng dụng; `hideGroupTools` đổi tên công cụ chuyên trong chuẩn làm việc thành tên nhóm (chuẩn làm việc PowerPoint từng nêu `ppt_add_chart`…); nhãn bước `load_tools` tra trong mọi nhóm (`ALL_TOOL_GROUPS`, cùng id thì cùng nhãn: nhãn nhóm `ngoai` của Word đổi thành "ứng dụng Office khác"); ảnh đính kèm trong PowerPoint bật nhóm `hinh_anh`.
- Từ khóa có dấu cho các chữ dễ lẫn khi bỏ dấu: "lọc" (không lẫn "Lộc"), "trùng" (không lẫn "trung bình"), "khóa" (không lẫn "khoa"), "ảnh" (không lẫn "anh"), "lỗi" (không lẫn "lời"), "bảng". Nhóm in của Excel chỉ dùng cụm ("in ra", "khổ giấy", "vùng in"…) để "in đậm", "in nghiêng" không bật.
- Kiểm thử: `test/unit/dot11.test.ts` (26 ca: nhóm có thật và không trùng, nhóm lõi, từ khóa đúng và không bật nhầm, kỹ năng kéo theo nhóm, lời nhắc Miễn phí không nêu tên công cụ ẩn, một lượt Miễn phí xin nhóm rồi chạy, chặn phình lại); M5 của Đợt 10 sửa theo. Bài đo `kiem-thu/do-token.ts --free` chạy cho cả ba ứng dụng. Bài thử `kiem-thu/cdp-bai/dot-11-excel-du-lieu.js`, `dot-11-ppt-du-lieu.js` (dựng lại tài liệu thử trước mỗi bài, không phải mở lại ứng dụng), `dot-11-chay.js`.
- Số bản: 4.11.0 (chưa phát hành).

## 6. Kết quả (20/11/2026, desktop ẩn)

### Phần cố định mỗi lượt (`kiem-thu/do-token.ts`, ký tự)

| Ứng dụng | Mốc 4.10.0 | Trả phí (A) | Miễn phí, nhóm lõi (A + B) |
|---|---|---|---|
| Excel | 31.754 (33 công cụ) | 25.417 (**−20,0%**) | 14.710, 19 công cụ (**−53,7%**) |
| PowerPoint | 23.293 (27 công cụ) | 18.745 (**−19,5%**) | 13.888, 17 công cụ (**−40,4%**) |
| Word (đối chiếu) | 30.049 sau Đợt 10 | 30.049 (không đổi) | 18.604 (nhãn nhóm `ngoai` ngắn hơn) |

### Chạy thật, Claude Haiku 5.5, Trả phí (cùng bài, cùng dữ liệu dựng lại, hội thoại mới mỗi bài)

"Trước" = bản 4.10.0 đã phát hành; "sau" = bản cục bộ 4.11. Token vào gồm phần đọc từ và ghi vào bộ nhớ đệm.

| Bài | Trước: lần gọi · token vào | Sau: lần gọi · token vào | Mỗi lần gọi | Kết quả |
|---|---|---|---|---|
| E1 Thêm cột giá bình quân bằng công thức | 3 · 49.120 | 3 · 41.769 | 16.373 → 13.923 (−15,0%) | Đúng cả hai; bản sau tự quét lỗi công thức |
| E2 Biểu đồ cột tổng doanh thu theo khu vực | 3 · 52.477; lần 2: 4 · 71.100 | 4 · 60.658; lần 2: 4 · 60.615 | ~17.600 → ~15.160 (−14%) | Đúng cả hai (bảng SUMIF + biểu đồ, cùng số liệu). Bản cũ cũng có lần 4 lượt gọi: số lần gọi là dao động của model |
| E3 PivotTable theo khu vực, dự án | 2 · 31.881 | 2 · 26.943 | 15.941 → 13.472 (−15,5%) | Đúng cả hai |
| P1 Thêm 2 slide | 2 · 25.783 | 3 · 35.562 (tự chụp 2 slide để kiểm); lần 2: 2 · 21.983 | 12.892 → 10.992 (−14,7%, lần 2) | Đúng cả hai |
| P2 Slide biểu đồ cột 3 năm | 2 · 27.517 | 4 · 52.301 (chụp, sửa chữ, chụp lại); lần 2: 2 · 25.056 (có chụp slide) | — (lần 2 có thêm ảnh slide) | Đúng cả hai |
| P3 Kiểm bài, chỉ báo lỗi | 2 · 26.358 | 2 · 25.930 (thêm chụp 4 slide) | — | Cả hai: `ppt_check_deck` đủ 4 điểm cài sẵn (2 phông lệch mẫu, 1 chữ cỡ 9, 1 chữ mẫu sót) |

Bản sau ở PowerPoint hay tự chụp slide để kiểm hơn (đúng dặn dò "TỰ KIỂM BẰNG MẮT" vốn có ở cả hai bản); ảnh slide tốn thêm token nên lần chạy có chụp không so được theo số mỗi lần gọi.

### Chế độ Miễn phí (bản cục bộ)

| Bài | Nguồn | Lần gọi · token vào | Kết quả |
|---|---|---|---|
| EF1 T4 Excel: "Chỉ hiện các dòng của khu vực Đà Nẵng thôi." (không khớp từ khóa) | Gemini im lặng → Groq đọc được (lần gọi đầu chạy) rồi hết lượt → OpenRouter Nemotron 3 Ultra free | 5 · 33.641 | Model gọi `load_tools(loc_dinh_dang)` → `excel_filter` cột B → đúng 6 dòng |
| EF2 Excel: biểu đồ cột theo khu vực (khớp từ khóa) | Gemini 3.8 Flash | 5 · 39.961 | Nhóm `bieu_do` bật sẵn; bảng SUMIF + biểu đồ đúng số |
| PF1 T4 PowerPoint: "Đưa slide cuối cùng lên làm slide thứ hai." (không khớp từ khóa) | Gemini im lặng → Groq hết lượt → OpenRouter Nemotron | 4 · 24.788 | `load_tools(sap_xep)` → `ppt_move_slide` 4 → 2; đọc lại bài: "Kiến nghị" ở slide 2 |
| PF2 PowerPoint: slide biểu đồ (khớp từ khóa) | Gemini 3.8 Flash | 3 · 24.764 | Nhóm `bieu_do_so_do` bật sẵn; biểu đồ gốc đúng 3 năm |

### Bảng nghiệm thu

| # | Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|---|
| T1 | Trả phí giảm ≥ 20%; Miễn phí nhóm lõi giảm ≥ 45% | Excel **Đạt** (−20,0% · −53,7%). PowerPoint **chưa đạt ngưỡng**: −19,5% · −40,4%. PowerPoint có ít công cụ chuyên (nhóm lõi giữ 17/27 công cụ, gồm chụp slide, đọc mẫu, định dạng mà lời nhắc dặn dùng thường xuyên); rút thêm thì phải ẩn công cụ tự kiểm, không nên | `do-token-*.json`; kiểm thử chặn phình lại |
| T2 | Token mỗi lần gọi Haiku giảm ≥ 15% | Excel **Đạt** (−15,0%, −15,5%; E2 −14%). PowerPoint **sát ngưỡng**: −14,7% (P1 cùng bước) | `truoc-*.json`, `sau-*.json` |
| T3 | Bài Haiku xong đúng, số lần gọi không tăng quá 1; mỗi ứng dụng một bài Miễn phí xong | **Đạt về kết quả**: 6/6 đúng, 4/4 bài Miễn phí xong. Số lần gọi: lần chạy đầu của P2 tăng 2 (tự chụp, sửa, chụp lại); lần chạy lại bằng bản cũ. E2 bản cũ cũng dao động 3–4 | như trên, `mien-phi-*.json` |
| T4 | Miễn phí, bài không khớp từ khóa → `load_tools` rồi xong | **Đạt** cả Excel và PowerPoint | `mien-phi-EF1.json`, `mien-phi-PF1.json` |
| T5 | Kiểm thử máy, kiểm kiểu; Tự kiểm tra | Kiểm thử máy 697/697, tsc sạch; Tự kiểm tra trên bản cục bộ: **Excel 40/40, PowerPoint 25/25**. **Word chưa chạy**: lúc tới lượt Word thì anh đang mở Excel (tệp của anh) trên desktop chính, quy định không mở Office thử lúc đó | `tu-kiem-tra-excel.json`, `tu-kiem-tra-powerpoint.json` |
| T6 | Soát mã độc lập | xem mục 7 | `soat-ma-dot-11.md` |
