# Vòng thử 3 — model miễn phí, đề mới, sửa add-in 4.11.11 (24/11/2026)

Tiếp theo `tasks/thu-mien-phi-vbhc.md` (bản 4.11.10, commit `e20dd9a`). Anh Dũng giao: tiếp tục vòng lặp thử nghiệm → hoàn thiện → thử lại với các đề mới để tối ưu add-in.

**Cách chạy** (như vòng trước): Word trên desktop ẩn, bản cục bộ nạp qua `cdp.mjs serve`, chế độ Miễn phí, mỗi bài một hội thoại mới; tài liệu trắng (bài sửa văn bản: tài liệu có sẵn đoạn mồi). Lần này bộ chạy được lưu vào kho để dùng lại: `kiem-thu/thu-mien-phi.mjs` + `kiem-thu/cdp-bai/thu-mien-phi-van-ban.js` (gom chat, định dạng từng đoạn, sửa đổi chờ duyệt, tệp .docx). Đề ở `tasks/thu-mien-phi-vong3-bang-chung/de.json`; bằng chứng `ban-4.11.10/` (bản cũ) và `ban-4.11.11/` (bản sửa).

## Đề mới (mỗi đề một chỗ phải tự suy luận hoặc bắt lỗi)

| Đề | Chỗ kiểm |
|---|---|
| Giấy mời UBND phường họp bàn bồi thường, "14h00 Thứ Tư, ngày 10/12/2026" | 10/12/2026 là Thứ Năm; ký KT. TỔNG GIÁM ĐỐC |
| Báo cáo tiến độ dự án gửi Sở Xây dựng | Tự tính 157,5/450 = 35%, 48/120 = 40% |
| Hợp đồng đặt cọc mua nhà hình thành trong tương lai, cọc 10% | Luật KDBĐS 2023: chủ đầu tư thu cọc tối đa 5% |
| Hợp đồng lao động kỹ sư, thử việc 90 ngày, lương thử việc 70% | Bộ luật Lao động 2019: tối đa 60 ngày, ít nhất 85% |
| Giấy ủy quyền TGĐ → Phó TGĐ ký hợp đồng mua bán nhà, hỏi có được ủy quyền lại không | Bộ luật Dân sự 2015: chỉ khi bên ủy quyền đồng ý; hai bên ký |
| Sửa văn bản có sẵn (công văn có lỗi) | "đề nghi", "dư án", "Trân trong"; Luật KDBĐS 2014, NĐ 99/2015 hết hiệu lực; 15/12/2026 là Thứ Ba (ghi Thứ Hai) |

Model: dots-3-note, Nemotron 3 Ultra (OpenRouter), Cohere Command A Plus.

## Điểm trên bản 4.11.10 (thang 10)

| Model | Giấy mời | Báo cáo | Đặt cọc | HĐLĐ | Ủy quyền | Sửa văn bản | TB |
|---|---|---|---|---|---|---|---|
| Nemotron 3 Ultra | 6 | 8 | 6 | 5 | 3 | 6 | **5,7** |
| Cohere A Plus | 6 | 8 | — * | — * | 6 | 2 | 5,5 (4 bài) |
| dots-3-note | 6 | 9 | 4 | 3 | 2 | 4 | 4,7 |

\* Cohere hết lượt giữa bài, add-in tự đổi sang Nemotron 3 Ultra (có dòng "Trả lời bởi …") → không chấm cho Cohere.

Thấy chung ở cả ba model:
- **Ngày ban hành giấy mời trùng ngày họp (10/12)** — 3/3 model; trích yếu giấy mời để "…" — 2/3.
- **Chép nguyên số trái luật**: cọc 10% (3/3 bài chấm được), thử việc 90 ngày + lương 70% (3/3). Không model nào cảnh báo.
- **Bên là cá nhân** (người lao động, người mua) vẫn ký "ĐẠI DIỆN BÊN B / <chức vụ> / (Ký, ghi rõ họ tên, đóng dấu)" — do công cụ.
- **Giấy ủy quyền không có loại riêng**: dots dựng thành công văn (V/v, Kính gửi, một bên ký), Nemotron viết thành quyết định HĐQT, Cohere dùng loại giấy mời.
- **Sửa văn bản**: không model nào gọi công cụ soát; dots tính 15/12/2026 là "Thứ Sáu", Cohere bỏ qua; cả ba bịa số hiệu nghị định thay thế (33/2023, 10/2024, 30/2021). Kỹ năng áp dụng lại là "chuẩn bị họp" — từ khóa "hop voi" khớp nhầm vào "k**hớp với**".
- **Cohere kể 5 chỗ đã sửa nhưng chỉ ghi 2** ("Trân trọng", thứ, căn cứ không hề sửa) — cảnh báo "nói đã sửa mà không ghi" không bắt vì lượt có ghi.
- dots chép "giấy\*uỷ\*quyền" (dấu \* dính chữ thay khoảng trắng); Cohere chép nguyên ví dụ kỹ năng "## BÊN GIAO THẦU (BÊN A):" không tên vào hợp đồng đặt cọc.

## Lỗi add-in đã sửa ở 4.11.11

| # | Lỗi | Sửa |
|---|---|---|
| B1 | Giấy mời ban hành trùng / sau ngày họp | `issueAfterEvent` (`textcheck.ts`) báo trong `check_now` |
| B2 | Giấy mời bỏ trống trích yếu | `check_now` nhắc viết trích yếu |
| B3 | Chép nguyên số trái luật (cọc > 5% của chủ đầu tư; thử việc > 60/180 ngày; lương thử việc < 85%; HĐLĐ xác định thời hạn > 36 tháng) | `src/core/legalcheck.ts` (mới), chạy trong `contentChecks` và trong công cụ soát; kỹ năng `soan-hop-dong` có mục "Ngưỡng luật hay gặp" |
| B4 | Bên cá nhân ký "ĐẠI DIỆN … đóng dấu" | `partyIsPerson`: bên cá nhân ký "BÊN B", không chức vụ, "(Ký, ghi rõ họ tên)" |
| B5 | Không có loại giấy ủy quyền | Loại `giay_uy_quyen`: tên "GIẤY ỦY QUYỀN", không Kính gửi, hai bên ký (trái: người nhận ủy quyền); nhận lại khi model dùng cong_van / khac / giay_moi với tên "GIẤY ỦY QUYỀN" |
| B6 | "giấy\*uỷ\*quyền" | `stripModelMarkup` đổi dấu \* dính chữ thành khoảng trắng |
| B7 | Công cụ soát (`word_check_consistency`) không bắt thứ sai, luật hết hiệu lực, sai dấu khi rà văn bản có sẵn | Loại điểm mới `fact` (mặc định chạy): thứ lệch ngày, văn bản hết hiệu lực (dặn không tự viết số hiệu), cụm sai dấu, từ lạ, ngưỡng luật; Tự kiểm tra Word cài thêm lỗi thứ lệch ngày |
| B8 | Cụm chuẩn sai dấu "đề nghi", "dư án", "Trân trong" lọt | Thêm vào danh sách cụm chuẩn của `accentTypos` |
| B9 | Kỹ năng tự kích hoạt khớp chuỗi con ("hop voi" trong "khop voi") | `autoSelectSkills` so theo ranh giới từ; kỹ năng văn bản hành chính thêm từ khóa "rà soát văn bản", "sửa lỗi chính tả", "giấy ủy quyền" và bước "gọi word_check_consistency TRƯỚC" |
| B10 | AI kể "A" → "B" đã sửa mà không ghi (lượt có ghi chỗ khác) | `unbackedEditClaims` (`claims.ts`): đối chiếu từng cặp với tham số các lệnh ghi thành công, cảnh báo trên khung (chỉ hiển thị, không vào lịch sử) |
| B11 | Ví dụ kỹ năng chép nguyên ("BÊN GIAO THẦU" không tên trong hợp đồng đặt cọc) | `check_now` báo dòng bên chưa có tên và vai trò giao/nhận thầu sai loại hợp đồng |
| Bộ thử | `busy()` coi nút "Dừng ở đây" là đang chạy; bảng chọn model mở sẵn thì bấm lại là đóng | Sửa `cdp-lib.js`; bộ chạy chờ khung rảnh trước mỗi bài |

Kiểm thử máy **869/869** (thêm `test/unit/vong3-mien-phi.test.ts`), kiểm kiểu sạch, phần cố định lượt Word Trả phí vẫn dưới trần 31.600 ký tự (rút gọn mô tả `word_check_consistency`), kỹ năng văn bản hành chính 10.997 < 11.000 ký tự. Tự kiểm tra Word trên bản cục bộ 4.11.11: **41/41** (`tu-kiem-tra-4.11.11.json`).

## Kiểm chứng trên 4.11.11 (cùng đề, cùng model)

| Model | Giấy mời | Báo cáo | Đặt cọc | HĐLĐ | Ủy quyền | Sửa văn bản | TB (trước → sau) |
|---|---|---|---|---|---|---|---|
| Nemotron 3 Ultra | 8 | 8 | 7 | 9 | 8 | 10 | **5,7 → 8,3** |
| Cohere A Plus | 8 | 8 | — * | — * | 8 | 6 | 5,5 → **7,5** (4 bài) |
| dots-3-note | 7 | 9 | 5 | 7 | 6 | 5 | 4,7 → **6,5** |

\* Cohere lại hết lượt giữa bài, Nemotron trả lời thay (đặt cọc: tự hạ cọc về 5% = 310 triệu; HĐLĐ: nêu rõ 90 ngày / 70% trái luật và hỏi).

Thấy rõ trên bản sửa (bằng chứng `ban-4.11.11/`):
- Giấy mời: 3/3 model không còn ghi ngày ban hành trùng ngày họp (bỏ trống, hoặc Cohere ghi 09/12), cả 3 có trích yếu.
- Ngưỡng luật: không model nào chép im lặng nữa — Nemotron sửa HĐLĐ về 60 ngày / 85% và nêu trần cọc 5%; dots, Cohere nêu rõ và hỏi người dùng; dots tự chèn nhận xét rủi ro vào điều tiền cọc.
- Bên cá nhân ký "BÊN B / (Ký, ghi rõ họ tên)" ở mọi hợp đồng.
- Giấy ủy quyền dựng đúng loại ở 3/3 model, hai bên ký; Nemotron và Cohere ghi đúng "chỉ được ủy quyền lại khi bên ủy quyền đồng ý bằng văn bản".
- Sửa văn bản: Nemotron và Cohere gọi `word_check_consistency` trước (5 điểm "sự kiện"), Nemotron sửa đủ 5 lỗi và không bịa số nghị định; Cohere sửa 4/5 và **cảnh báo mới (B10) bật đúng** khi Cohere kể đã sửa "Trân trong" mà không ghi.
- Còn lại: dots-3-note vẫn yếu ở bài sửa (không gọi công cụ soát, chèn "Denied" vào đoạn, bỏ sót "đề nghi") và vẫn lọt "\*uỷ\*" sau dấu phẩy — hai lỗi này là lý do có thêm B12, B13 bên dưới, kiểm lại ở mục "Bản cuối".

## Soát mã độc lập (4.11.10 + 4.11.11)

Một tác tử soát độc lập toàn bộ mã từ `f9015a8` (gồm 4.11.10 chưa soát): 3 lỗi cao, 8 trung bình, 4 thấp — phần lớn là **báo nhầm** của các phép kiểm mới. Đã sửa hết mức cao và trung bình, kèm ca kiểm thử "không được báo" cho từng mục (`test/unit/vong3-mien-phi.test.ts`, khối "Ca dễ báo nhầm"):

| # | Lỗi | Sửa |
|---|---|---|
| 1 | "Lương thử việc 100%" đọc thành "00%" | Bắt số không cắt giữa chừng, nhận 3 chữ số |
| 2 | "Phiên họp thứ 2, ngày …", "lần thứ hai" bị coi là thứ trong tuần | Bỏ khi đứng sau lần, kỳ, phiên, đợt…; "thứ" bằng số chỉ nhận trong ngoặc sau ngày |
| 3 | "Bên A có trách nhiệm: …", "SAMSUNG ELECTRONICS", "Kho bạc" bị coi là bên cá nhân | Chỉ xét phần giới thiệu các bên (trước Điều 1), bỏ dòng có động từ, họ tên phải là 2–5 âm tiết tiếng Việt, mở rộng danh sách chữ của tổ chức |
| 4 | "phạt cọc 50%", cọc giữa hai cá nhân bị báo trần 5% | Bỏ mức phạt; phải là tỷ lệ của giá; phải có bên là doanh nghiệp |
| 5 | Phụ cấp sau lương thử việc bị đọc là lương thử việc | Tỷ lệ phải gắn với lương, bỏ khi có phụ cấp, thưởng ở giữa |
| 6 | Dòng căn cứ, số văn bản bị coi là giờ họp | Bỏ dòng căn cứ; giờ chặn bằng `(?!\p{L})` thay `\b` |
| 7 | `stripModelMarkup` làm hỏng chữ đậm hợp lệ "**Ông A**, Chức vụ" | Chỉ gỡ khi một từ kẹp giữa hai dấu \* dính liền chữ |
| 8 | "có phần", "cổ động", "Trần Trọng" bị báo sai dấu | Bảng cụm đúng chính tả |
| 9 | Email, trang web, hợp đồng song ngữ bị báo từ lạ | Bỏ email, URL; bỏ đoạn ngoại ngữ; tối đa 8 điểm |
| 10 | Cảnh báo B10 bật khi model chép danh sách đề xuất "chưa sửa", hoặc cặp có "…" | Chỉ xét khi có "đã sửa / đổi…", bỏ dòng đề xuất, bỏ "…" khi so |
| 11 | Công văn có Kính gửi "V/v ủy quyền giao dịch ngân hàng" bị ép thành giấy ủy quyền | Công văn chỉ đổi loại khi không có Kính gửi |
| 12 (thấp) | Hợp đồng thuê nhắc "hợp đồng lao động", thuê 60 tháng bị báo trần 36 tháng | Chỉ xét câu "xác định thời hạn" |
| 13 (thấp) | "Nếu công cụ gặp lỗi…", "Công cụ không gặp lỗi" bị coi là đổ lỗi công cụ (có từ 4.11.10) | Bỏ câu giả định, phủ định |
| 15 (một phần) | Họ tên bên B viết hoa không được đổi | `personName` cho cả bên B |

Chưa làm (thấp, ghi lại): `boldLead` chỉ đậm "Điều N." cả với hợp đồng có tên điều dài; `consistency.ts` nhập `adminDoc.ts` chỉ để lấy `accentTypos`; trích yếu tách từ title bị viết thường cả tên riêng; `expiredLaws` khuyên thay căn cứ cả với hợp đồng cũ đã ký đúng luật lúc đó; `wordAdmin.ts:355` gán chữ đậm sau `getFirstOrNullObject()` chưa kiểm `isNullObject`; phép bắt chữ mẫu `<…>` báo nhầm "tỷ lệ < 50% và > 30%".

## Sửa thêm sau lượt kiểm chứng

| # | Lỗi | Sửa |
|---|---|---|
| B12 | Model chèn từ tiếng Anh lạc khi SỬA đoạn ("DeniedĐến nay", "Collaboratively配合", "bringing to") — kiểm từ lạ chỉ chạy lúc dựng văn bản mới | `newTextCheck` (`textcheck.ts`): kết quả `word_replace_text`, `word_set_paragraph_text`, `word_insert_paragraphs`, `word_insert_clause` có `check_now` khi chữ mới có từ lạ mà chữ cũ không có; từ Latinh viết hoa chỉ tính khi dính liền chữ Việt (không báo "Riverside") |
| B13 | Model tự "sửa" "Unauthorized" thành "Uy ban" (mất dấu) | Thêm cụm "Ủy ban nhân dân" vào kiểm sai dấu; hai kiểu đặt dấu "Ủy"/"Uỷ" coi như nhau |

Kiểm thử máy **885/885**, kiểm kiểu sạch. Tự kiểm tra Word trên bản cuối: **41/41** (`tu-kiem-tra-4.11.11.json`).

## Bản cuối (sau soát mã và B12, B13)

Chạy lại 3 bài dots-3-note còn yếu nhất (`ban-4.11.11-cuoi/`):

| Bài | 4.11.10 | 4.11.11 | Bản cuối | Ghi chú |
|---|---|---|---|---|
| Sửa văn bản | 4 | 5 | 6 | Không còn chèn từ lạ; sửa 4/5 (sót "đề nghi"); vẫn tự viết số hiệu "Nghị định 100/2024" dù kỹ năng và công cụ đã dặn; không gọi công cụ soát |
| Giấy ủy quyền | 2 | 6 | 7 | Hết dấu \* dính chữ; ghi rõ không được ủy quyền lại; hai bên ký; trích yếu để "…" → thêm B14 |
| Đặt cọc | 4 | 5 | 6 | Tự sửa "shall", "当" theo `check_now`; ghi nhận xét rủi ro cọc 10% > 5% |

B14 (sau lượt này): giấy ủy quyền bỏ trống trích yếu cũng được nhắc trong `check_now` như giấy mời. Kiểm thử máy **886/886**; B14 là hàm thuần, chưa Chạy thật lại.

## Chưa làm, để lại

- **Chưa phát hành**: nhánh `claude/thu-mien-phi-vong-3` (tách từ `e20dd9a`, tức 4.11.10 chưa phát hành) — phát hành là gộp vào `main` rồi đẩy.
- dots-3-note: khi rà văn bản vẫn không gọi `word_check_consistency` và vẫn tự viết số hiệu nghị định. Có thể cân nhắc: khi yêu cầu có "rà soát / sửa lỗi" ở chế độ Miễn phí thì add-in tự chạy soát bằng mã và đưa kết quả vào lượt đầu (không chờ model gọi).
- Model hay tự đặt "Hà Nội" làm địa danh khi đề không cho (Nemotron, Cohere).
- Các mục mức thấp của bản soát mã (liệt kê ở trên).
- Cohere hết lượt giữa bài ở cả hai lượt chạy (đặt cọc, HĐLĐ) → chưa có điểm Cohere cho hai đề hợp đồng.

## Gợi ý chọn model miễn phí cho Word (cập nhật)

1. **Nemotron 3 Ultra**: 8,3 trên 4.11.11, tốt nhất ở cả soạn mới lẫn sửa văn bản có sẵn (gọi công cụ soát, không bịa số hiệu).
2. **Cohere A Plus**: 7,5 (4 bài), nhanh; hay hết lượt giữa chừng → để làm nguồn dự phòng.
3. **dots-3-note**: 6,5; soạn mới ổn, sửa văn bản có sẵn còn yếu.
