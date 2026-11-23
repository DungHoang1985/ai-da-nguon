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
| Kiểm thử máy, kiểm kiểu | Đạt 814/814, kiểm kiểu sạch | `npm test` |
| Tự kiểm tra Word (bản cục bộ, desktop ẩn) | Đạt 41/41 | `sua-23-11-bang-chung/tu-kiem-tra-word-41.json` |

Chưa làm: soát mã độc lập; phát hành (đẩy `main`) — chờ anh duyệt; Chạy thật với Claude (Trả phí).
