# Bản vá 4.11.8 — soạn hợp đồng đúng thể thức (23/11/2026)

Anh Dũng báo: dùng chế độ Miễn phí (Nemotron 3 Ultra free) soạn hợp đồng xây dựng thì nội dung tốt nhưng thể thức sai — thừa dòng "HỢP ĐỒNG XÂY DỰNG"; nhờ xóa bớt dòng đó thì phần đầu (tên đơn vị trái, quốc hiệu phải) bị chỉnh lung tung. Bản chính thức (Sonnet, Opus) làm đúng.

## Nguyên nhân (đọc lại hội thoại đã lưu trên máy, chỉ phần tin nhắn và tên công cụ — không đọc, không in key)

1. **Do add-in**: công cụ `word_create_admin_document` không có loại hợp đồng. Model dùng `khac` + title "HỢP ĐỒNG XÂY DỰNG", đồng thời chép lại tên loại, "Số:", ngày vào đầu body → hai dòng tên; công cụ còn thêm Kính gửi / Nơi nhận / "./." như công văn.
2. **Do add-in**: phần đầu là bảng hai cột không viền, nhưng `word_read_paragraphs` không cho biết đoạn nào trong ô bảng. Model đọc chỉ số đoạn (ô trái hết rồi mới tới ô phải) thành thứ tự trên dưới, kết luận "tên cơ quan đứng trước quốc hiệu là sai", xóa tên đơn vị rồi chèn lại dưới tiêu ngữ → lệch phần đầu.
3. **Do model** (thấy khi Chạy thật): Nemotron tự ghi số hiệu luật từ trí nhớ (có văn bản không tồn tại), đôi khi lẫn chữ Hán / Ả Rập vào nội dung.

## Đã sửa

| Chỗ | Sửa |
|---|---|
| `src/core/adminDoc.ts` | Loại `hop_dong` (nhận cả "hợp đồng…", "hop_dong_xay_dung"): có `org` → bảng hai cột (trái: đơn vị, "Số:"; phải: quốc hiệu, tiêu ngữ, ngày nếu có); không `org` → quốc hiệu canh giữa, "Số:" dưới tên; tên hợp đồng một dòng; không Kính gửi / Nơi nhận / "./."; phần ký ĐẠI DIỆN BÊN A / BÊN B cùng hàng; dòng "Căn cứ" tự in nghiêng; số hợp đồng đủ dạng giữ nguyên. `stripHeaderEcho`: bỏ dòng lặp phần đầu ở ĐẦU body và phần ký ở CUỐI body (mọi loại văn bản). `foreignRuns`: báo cụm chữ lạ. |
| `src/office/wordAdmin.ts` | Mô tả công cụ, `org` không bắt buộc, `left_signer_authority`; kết quả có `foreign_text` + `fix_now` khi lẫn chữ lạ; lời dặn sửa phần đầu; bỏ đoạn rỗng Word tự sinh sau bảng cuối (đo trong Word: phần ký hợp đồng sinh một đoạn "Normal" rỗng, Hoàn tác để sót). |
| `src/office/wordTools.ts`, `util.ts` | `word_read_paragraphs` gắn `cell` ("bảng 1, hàng 1, cột 2") cho đoạn trong ô và `table_note` (hai ô cùng hàng nằm NGANG); `word_delete_paragraphs` từ chối vùng vắt qua mép bảng. |
| `src/core/skillpack.ts` | Kỹ năng mới `soan-hop-dong` (20 kỹ năng); kỹ năng hành chính chỉ hợp đồng sang `hop_dong`, dặn body chỉ là nội dung, dặn không "sắp lại" phần đầu. |
| `src/core/toolgroups.ts` | Nhóm `hanh_chinh` thêm từ khóa soạn hợp đồng, quốc hiệu, tiêu ngữ và kỹ năng `soan-hop-dong`. |
| `src/office/selftestWord.ts` | Bước Tự kiểm tra "Hợp đồng" (41 bước). |
| `test/unit/hop-dong.test.ts` | 13 kiểm thử; trần token Trả phí 31.100 → 31.600 (+426 thật). |

## Nghiệm thu

| Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|
| Hợp đồng có tên đơn vị: trái đơn vị, phải quốc hiệu; tên hợp đồng một dòng; không Nơi nhận / "./." | Đạt | Kiểm thử máy; Tự kiểm tra Word bước "Hợp đồng"; Chạy thật Nemotron 3 lần |
| Body chép lại phần đầu / phần ký thì tự bỏ | Đạt | Kiểm thử máy; Tự kiểm tra (body có dòng lặp, công cụ ghi chú "đã bỏ") |
| Đọc đoạn cho biết ô bảng, hàng, cột | Đạt | Tự kiểm tra: "bảng 1, hàng 1, cột 1 / cột 2" |
| Xóa vắt mép bảng bị chặn, văn bản không đổi | Đạt | Tự kiểm tra |
| Hoàn tác gỡ sạch hợp đồng | Đạt | Tự kiểm tra (sau khi bỏ đoạn rỗng tự sinh) |
| Câu "đang thừa một dòng Hợp đồng xây dựng": chỉ xóa đúng dòng, phần đầu giữ nguyên | Đạt | Chạy thật Nemotron: xóa đúng đoạn, 8 đoạn phần đầu vẫn trong bảng như cũ |
| Soạn mẫu không hỏi dồn | Đạt | Chạy thật lần 1 (trước sửa kỹ năng) hỏi 7 mục; lần 2–4 soạn ngay |
| Căn cứ in nghiêng; không tự ghi số hiệu luật | Đạt | Chạy thật lần 3: căn cứ nghiêng, số hiệu để "…" |
| Chữ lạ trong nội dung được báo để model sửa | Kiểm thử máy đạt; Chạy thật chưa kiểm được (lần 4 model không sinh chữ lạ) | `hop-dong.test.ts` |
| Kiểm thử máy, kiểm kiểu | Đạt 819/819 (sau soát mã), kiểm kiểu sạch | `npm test` |
| Tự kiểm tra Word (bản cục bộ, desktop ẩn) | Đạt 41/41 | `sua-23-11-bang-chung/tu-kiem-tra-word-41.json` |

## Soát mã độc lập (23/11/2026, anh duyệt "soát mã rồi phát hành")

| Mức | Phát hiện | Xử lý |
|---|---|---|
| Cao C1 | Dòng lặp bị bỏ mà không giữ giá trị: model dồn tên, số, ngày vào body thì văn bản mất tên, số, ngày | Đã sửa: `stripHeaderEcho` trả `found` (tên, số, địa danh ngày), bên gọi dùng khi tham số trống (hợp đồng và văn bản hành chính) |
| Cao C2 | Từ khóa "mẫu hợp đồng" bật kỹ năng soạn khi người dùng nhờ rà soát, góp ý, tóm tắt, điền mẫu | Đã sửa: chỉ cụm có động từ soạn ("soạn mẫu hợp đồng"…); thêm "rà soát / kiểm tra / góp ý mẫu hợp đồng" cho ra-soat-hop-dong, "điền mẫu hợp đồng" cho dien-mau; bỏ khỏi nhóm hanh_chinh |
| TB T1 | Quy tắc "Số" bỏ nhầm "Số liệu …", "So sánh 2024/2025" | Đã sửa: phải có "Số:" (còn dấu, có hai chấm) |
| TB T2 | Bỏ đoạn rỗng sau bảng ký có thể xóa đoạn ngăn với bảng có sẵn → hai bảng dính | Đã sửa: chỉ bỏ khi ngay sau nó là đoạn vốn đứng sau chỗ chèn; đoạn ngăn được giữ và Hoàn tác gỡ ngay sau bảng ký (Tự kiểm tra thêm ca chèn sát bảng — lần đầu thấy Hoàn tác sót đoạn ngăn, đã sửa) |
| TB T3 | "Căn cứ …" trong điều khoản cũng bị nghiêng | Đã sửa: chỉ dòng căn cứ ở phần mở đầu |
| TB T4 | Chặn xóa bỏ lọt vùng vắt từ ô trái sang ô phải; bị chặn sau khi người dùng đã bấm đồng ý | Đã sửa: chặn cả vùng nhiều ô; kiểm ngay ở bước hỏi (không hỏi rồi mới từ chối); mô tả công cụ ghi giới hạn; câu lỗi nói cách xóa từng phần |
| Thấp L1 | Bảng con gắn nhãn của bảng ngoài | Đã sửa: ghi "(bảng con trong ô)" |
| Thấp L2 | Phần ký có họ tên, nơi nhận có danh sách ở cuối body không bị bỏ; câu "HỢP ĐỒNG NÀY …" bị bỏ nhầm | Đã sửa, có kiểm thử |
| Thấp L3 | Thêm lượt sync cho mọi văn bản; chèn cuối tài liệu không dọn đoạn thừa | Sửa phần đầu (chỉ nạp khi khối cuối là bảng); ca chèn cuối tài liệu để lại (chưa đo trong Word) |
| Thấp L4, L5, L7 | Mô tả cỡ chữ, org; CLAUDE.md lệch trần ký tự | Đã sửa |
| Thấp L6 | `foreign_text` lẫn chữ đứng trước làm model dùng sai `find`; cắt chữ theo đơn vị mã | Đã sửa: dạng `<cụm lạ> (sau “…”)`, cắt theo ký tự |

Sau soát mã: 819/819, Tự kiểm tra Word 41/41 (bản cục bộ, desktop ẩn), Chạy thật Nemotron đạt (bố cục đúng, không chữ lạ, căn cứ nghiêng đúng chỗ).

Chưa làm: Chạy thật với Claude (Trả phí); đo ca chèn hợp đồng vào cuối tài liệu đang có chữ.
