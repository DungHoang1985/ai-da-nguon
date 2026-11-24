# Thử model miễn phí soạn văn bản hành chính, sửa add-in 4.11.10 (23–24/11/2026)

Tiếp theo `tasks/viec-le/thu-mien-phi-hdcn/thu-mien-phi-hdcn.md` (hợp đồng chuyển nhượng QSDĐ, lỗi A1–A7). Anh Dũng giao: soạn đề theo tình huống giả định cho các loại văn bản, thử các model miễn phí (OpenRouter thử rộng, hạn mức 1.000 lượt/ngày), chấm điểm, rồi sửa lỗi add-in một lần.

Cách chạy: Word trên desktop ẩn, chế độ Miễn phí, mỗi bài một hội thoại mới trên tài liệu trắng; nếu lượt 1 chưa ghi vào Word thì nhắc một câu "cứ soạn luôn". Bằng chứng (.docx, hội thoại, định dạng từng đoạn) ở `tasks/viec-le/thu-mien-phi-vbhc/bang-chung/`: `ban-phat-hanh/` (4.11.9), `ban-sua-4.11.10/` (bản cục bộ đã sửa), `de.json` (đề bài).

## Đề bài (một công ty giả định, mỗi đề có một chỗ phải tự suy luận)

| Đề | Chỗ kiểm suy luận |
|---|---|
| Công văn gửi Sở Xây dựng đề nghị thông báo nhà ở hình thành trong tương lai đủ điều kiện bán | Căn cứ là Luật Kinh doanh BĐS 2023, không dẫn luật 2014 |
| Thông báo nghỉ Tết Dương lịch 2027, ghi rõ thứ | 01/01/2027 là Thứ Sáu → nghỉ 3 ngày, đi làm Thứ Hai 04/01; ký "KT. TỔNG GIÁM ĐỐC / PHÓ TỔNG GIÁM ĐỐC" |
| Tờ trình HĐQT vay 150 tỷ, tổng tài sản 380 tỷ, hỏi thẩm quyền | 150/380 = 39,47% ≥ 35% → HĐQT thông qua (điểm h khoản 2 Điều 153 Luật Doanh nghiệp 2020) |
| Quyết định HĐQT bổ nhiệm Phó Tổng Giám đốc | Ký "TM. HỘI ĐỒNG QUẢN TRỊ / CHỦ TỊCH"; không Kính gửi; Nơi nhận "Như Điều …" |
| Nghị quyết HĐQT 3 nội dung, số 15/2026/NQ-HĐQT | Add-in chưa có loại nghị quyết; "QUYẾT NGHỊ:" |
| Biên bản họp ĐHĐCĐ bất thường, cho số cổ phần | Dự họp 73% > 50% (Điều 145); tán thành 85% / 10% / 5%, ≥ 65% (Điều 148) |

## Model đã thử

- Chạy đủ 6 đề: Nemotron 3 Ultra, Nemotron 3.5 Super VL (bài "Gemini 3.8 Flash" thực ra do Nemotron 3.5 Super trả lời vì Gemini quá tải), Cohere Command A Plus, dots-3-note, openrouter/free (bộ tự chọn model).
- Vòng lọc 2 đề (Quyết định, Biên bản): Nemotron 3.5 Lightning, Nemotron 3 Nano Omni, Flint 1.0, LFM 3 1.7B. Ling 3.2 Flash, Particle-1, Gemma 4 26B và 31B, Gemini 3.5 Flash, Gemini flash-latest: hết lượt hoặc quá tải, nguồn khác trả lời thay → không chấm.
- Không có model nào tra cứu được: với nguồn miễn phí, add-in không có công cụ tra web khi AI tự thao tác.

## Điểm trên bản phát hành 4.11.9 (thang 10)

| Model | Công văn | Thông báo | Tờ trình | Quyết định | Nghị quyết | Biên bản ĐHĐCĐ | TB |
|---|---|---|---|---|---|---|---|
| dots-3-note | 5 | 7 | 3 | 7 | 3 | 8 | **5,5** |
| Nemotron 3.5 Super | 5 | 3 | 6 | 5 | 5 | 8 | 5,3 |
| openrouter/free | 6 | 3 | 7 | 6 | 1 | 9 | 5,3 |
| Nemotron 3 Ultra | 5 | 8 | 5 | 7 | 1 | 1 | 4,5 |
| Cohere Command A Plus | 4 | 2 | 5 | 4 | 2 | 6 | 3,8 |
| Flint 1.0 · Nano · Lightning · LFM 1.7B | — | — | — | 4 · 3 · 3 · 1 | — | 5 · 3 · 3 · — | |

Điểm thấp phần lớn do add-in, không do model: Nemotron 3 Ultra mất trọn Nghị quyết và Biên bản vì gửi `body` dạng mảng JSON (cả văn bản thành một đoạn `["…"]`); Nghị quyết của dots-3-note và openrouter/free không dùng được công cụ dựng vì từ "nghị quyết" không nạp nhóm công cụ hành chính ở chế độ Miễn phí.

Lỗi của model (add-in không sửa được, chỉ báo): Cohere và Nemotron 3.5 Super tính sai thứ trong tuần; Flint, Nano tính tỷ lệ 84,81%, 84,9% (đúng 85%); dots-3-note dẫn sai điều luật ở tờ trình ("Điều 167/168, 10%"); Nemotron 3 Ultra bịa ngân hàng "BIDV", tỷ lệ đặt cọc, số nghị định; Lightning, Flint, LFM sai chính tả nặng ("BIÊN BÁN HỘP", "CHỦ TỬC", "CỔ PHẪN").

## Kiểm chứng trên bản sửa 4.11.10 (bản cục bộ, cùng đề)

| Model | Công văn | Thông báo | Tờ trình | Quyết định | Nghị quyết | Biên bản | TB | Hợp đồng (đề cũ) |
|---|---|---|---|---|---|---|---|---|
| Nemotron 3 Ultra | 6 | 9 | 9 | 9 | 9 | 8 | **8,3** (từ 4,5) | 8 (từ 7) |
| dots-3-note | 7 | 8 | 8 | 8 | 6 | 9 | **7,7** (từ 5,5) | 6 |
| Cohere Command A Plus | 4 | 3 → 6* | 7 | 8 | 8 | 5 | **5,8** (từ 3,8) | 6 (từ 4) |

\* Lần chạy lại trên bản cuối: công cụ báo "5 chỗ nghi sai cần kiểm", Cohere sửa hết thứ trong tuần; ngày nghỉ vẫn sai (nghỉ tới 05/01) nhưng giờ ghi đúng thứ nên lộ rõ.

Thấy rõ trên bản sửa: ký đúng "KT. TỔNG GIÁM ĐỐC", "TM. HỘI ĐỒNG QUẢN TRỊ / CHỦ TỊCH" ở cả 3 model; tờ trình tự tính 39,47% và dẫn đúng điểm h khoản 2 Điều 153 (Nemotron, dots, Cohere); biên bản tính đúng 73%, 85/10/5%, ngưỡng 65%, dẫn Điều 145, 148; hợp đồng dẫn Luật Nhà ở 2023, Luật Đất đai 2024, Luật KDBĐS 2023 (trước: luật 2013–2014, Nghị định 99/2015); model tự sửa từ lạ theo `check_now` ("distraint", "bringing to", "opinions", chữ Hán). Còn: Nemotron vẫn bịa một số hiệu nghị định ở công văn; Cohere vẫn dẫn Nghị định 99/2015 ở công văn (bản cuối đã có kiểm bằng mã báo văn bản hết hiệu lực), tự đặt giờ họp 9h30 (đề cho 8h30), chèn thêm đoạn sau Nơi nhận.

Tự kiểm tra Word trên bản cục bộ 4.11.10: **41/41** (`tu-kiem-tra.json`). Kiểm thử máy **850/850**, kiểm kiểu sạch.

## Lỗi add-in đã sửa (A1–A7 của buổi hợp đồng, A8 trở đi của buổi này)

| # | Lỗi | Sửa (4.11.10) |
|---|---|---|
| A1 | Tự đổi nguồn mà chân khung vẫn ghi model đã chọn | Dòng "Trả lời bởi … (đã tự đổi nguồn)" dưới câu trả lời (`Chat.tsx`) |
| A1b | Model đang nghỉ (lỗi ở lượt trước) bị bỏ qua **lặng lẽ**, nguồn khác trả lời thay mà không có dòng "↪" (lộ ở vòng lọc: Ling, Particle, Gemma → Cohere) | `rotation.ts`: bỏ qua nguồn đang nghỉ cũng báo đổi nguồn kèm lý do và giờ dùng lại |
| A2 | Phần ký hợp đồng lặp "ĐẠI DIỆN BÊN A" hoặc lộn A/B | Chức vụ trùng nhãn bên thì bỏ; quyền hạn ghi nhầm bên dùng mặc định |
| A3 | Kiểm chữ lạ bỏ sót từ Latinh ("burden", "competencia", "cấmélé", "quórum", "chủthể") | `textcheck.ts`: kiểm âm tiết tiếng Việt cho từ viết thường; trả `check_now` để model sửa ngay |
| A4 | "## Điều 8" ghi nguyên chữ khi chèn điều khoản | `word_insert_clause` bỏ ký hiệu markdown |
| A5 | Hợp đồng dán từ chat có hai phần ký và khối "Lưu ý" | Bỏ phần ký tự viết ("Bên chuyển nhượng", "(Ký, ghi rõ họ tên)") và khối "Lưu ý" cuối, có ghi chú |
| A6 | Model nói "công cụ đang lỗi" mà không gọi công cụ; nói "sẵn sàng ký" khi còn chỗ trống | Cảnh báo mới `toolFailureClaim`; lời dặn trong kết quả công cụ và kỹ năng |
| A7 | Căn cứ pháp lý hết hiệu lực | Kỹ năng có danh mục căn cứ đang hiệu lực / đã hết hiệu lực (tên và năm, theo quy tắc không ghi số hiệu); mã kiểm báo khi văn bản dẫn luật 2013–2014, Nghị định 99/2015, 02/2022 |
| A8 | `body` là chuỗi mảng JSON → cả văn bản một đoạn | `bodyLinesOf` nhận mảng, chuỗi mảng JSON, "\n" viết bằng chữ |
| A9 | Thẻ trích dẫn Cohere `<co>…</co: 2:[0]>` lọt vào văn bản | Bỏ ở mọi tham số |
| A10 | Chưa có loại nghị quyết; "nghị quyết" không nạp nhóm công cụ ở chế độ Miễn phí | Loại `nghi_quyet` (nhận cả khac/quyet_dinh có tên NGHỊ QUYẾT), "QUYẾT ĐỊNH:" → "QUYẾT NGHỊ:"; từ khóa "nghị quyết", "biên bản họp", "đại hội đồng cổ đông" |
| A11 | Quyền hạn ký sai dạng ("TGĐ", "ĐẠI DIỆN PHÁP LUẬT", "TM." trơn, "KÝ THAY …", "KT. PHÓ …", trùng chức vụ) | `normalizeSigner`: đổi dạng đủ chữ ra TM./KT./TL./TUQ., bỏ dạng sai, báo "KT. PHÓ …"; nghị quyết thiếu TM. thì báo; tờ trình ký "TM." nơi nhận thì báo |
| A12 | Thứ trong tuần sai so với ngày | `weekdayIssues` báo trong `check_now`; dòng tóm tắt trên khung ghi "N chỗ nghi sai cần kiểm" |
| A13 | Kính gửi chép lại ở đầu nội dung; quyết định, nghị quyết, biên bản có Kính gửi | Bỏ dòng lặp (giữ giá trị làm Kính gửi); ba loại kia chuyển sang Nơi nhận |
| A14 | "## Điều 1. <lời văn dài>" in đậm cả đoạn | Chỉ "Điều 1." đậm (`boldLead`) |
| A15 | Số "15/2026/NQ-HĐQT" mất năm | Số đủ dạng có năm giữ nguyên |
| A16 | "Trân trọng,./.", "Như Điều 3.;", "Điều lệ công ty.;" | Chuẩn hóa dấu câu |
| A17 | Dòng tên loại kèm trích yếu lặp ở đầu nội dung; title "NGHỊ QUYẾT VỀ VIỆC …" | Bỏ dòng lặp, phần sau làm trích yếu |
| A18 | "_ " làm gạch đầu dòng; một cặp `_…_` bao nhiều dòng; mở `_` không đóng | Đổi thành "- "; nghiêng từng dòng; đóng lại |
| A19 | Chữ mẫu "THẨM QUYỀN BAN HÀNH" chép nguyên; tên đơn vị, cụm chuẩn sai dấu ("HỘI ĐÒNG", "CỔ PHẪN", "Hộp đại hội đồng", "NGHỊ SỨ") | Báo trong `check_now`; kỹ năng đổi cách viết ví dụ |
| A20 | Tờ trình có dòng "QUYẾT ĐỊNH:" | Báo trong `check_now` |
| A21 | Họ tên người ký toàn chữ hoa | Đổi về chữ thường (in đậm theo thể thức) |

Kỹ năng `soan-thao-van-ban-hanh-chinh` thêm nghị quyết, biên bản họp ĐHĐCĐ (Điều 145, 148), thẩm quyền vay/đầu tư (Điều 153, 138), quyền hạn ký, Kính gửi theo loại; trần độ dài kiểm thử của kỹ năng nâng 9.000 → 11.000 ký tự (kỹ năng chỉ nạp khi dùng). Phần cố định mỗi lượt Word Trả phí vẫn dưới trần 31.600 ký tự.

## Chưa làm, để lại

- **Chưa soát mã độc lập, chưa phát hành** (đang ở nhánh `claude/thu-mien-phi-van-ban`, chưa commit).
- Với model yếu, `check_now` có thể làm tốn bước: chạy lại trên bản cuối, Cohere được báo đúng "Nghị định 99/2015 hết hiệu lực" (công văn) và 5 chỗ sai thứ (thông báo), nhưng loay hoay tới trần 12 bước; công văn vẫn còn căn cứ cũ. Có thể cân nhắc: chỉ báo `check_now` một lần cho người dùng (dòng tóm tắt đã có) thay vì yêu cầu model sửa ngay khi model là nguồn hay sai.
- Model bịa số hiệu văn bản (Nemotron: "Nghị định 144/2024") — mã chỉ bắt được văn bản đã hết hiệu lực trong danh mục, không bắt văn bản không tồn tại.
- Lỗi suy luận của model (ngày nghỉ, giờ họp, điều luật) chỉ lộ nhờ báo thứ / tỷ lệ; add-in không tự sửa.
- Gemini cả ngày 24/11 quá tải → chưa có điểm Gemini cho văn bản hành chính.
- Lúc kiểm chứng, Word thử một lần tự thoát, lần sau mở hai bản cùng một tệp (bản thứ hai chỉ đọc → mọi lệnh ghi báo `GeneralException`). Mở tệp thử mới thì chạy đúng; không thấy sự kiện sập trong nhật ký Windows. Nghi do môi trường thử, chưa tìm ra nguyên nhân.
- Bộ thử `cdp-lib.js` coi khung đang chờ "Làm tiếp?" là còn bận (nút "Dừng ở đây" khớp `/^Dừng/`) → bài chạm trần 12 bước bị chờ tới hết giờ.

## Gợi ý chọn model miễn phí cho Word

1. **Nemotron 3 Ultra** (OpenRouter): trên 4.11.10 tốt nhất (8,3), suy luận pháp lý đúng, tự sửa chữ lạ; chậm hơn (40–100 giây/bài).
2. **dots-3-note** (OpenRouter): 7,7, nhanh, gọn, ít bịa.
3. Cohere Command A Plus: nhanh nhất (10–40 giây) nhưng hay sai ngày, giờ, căn cứ; dùng làm nguồn dự phòng.
4. Không nên: openrouter/free (mỗi lần một model, chất lượng nhảy), Lightning, Nano, Flint, LFM 1.7B.
