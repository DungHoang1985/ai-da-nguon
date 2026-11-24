# Soát mã độc lập — Đợt 1A (`git diff 217d4c4..HEAD -- src test`)

Phạm vi đọc: `src/core/outline.ts`, `src/core/worddiff.ts`, `src/office/wordTools.ts`, `src/office/util.ts`, `src/office/selftestWord.ts` (phần Hoàn tác và dọn), `test/unit/dot1a.test.ts`, `src/core/agent/loop.ts` (chỗ gọi công cụ).
Cách kiểm: đọc mã, và chạy thử các hàm thuần bằng ts-node (tệp thử đặt ngoài kho, không sửa kho). Mọi mục dưới đây đã lần theo mã; mục nào chỉ là nghi ngờ thì ghi "chưa chắc".

Kết quả chạy thử đáng chú ý:
- Thử ngẫu nhiên 200.000 cặp (cũ, mới) với bảng chữ nhỏ (a, b, ab, khoảng trắng, dấu, số): `applyHunks(old, diffWords(old,new)) === new` luôn đúng. Mô phỏng thi hành `planParagraphEdit` (kể cả có khúc chữ đã xóa xen giữa, spans đúng) cho ra đúng câu mới ở mọi trường hợp lập được kế hoạch (979 trường hợp trả lỗi chuỗi, không có trường hợp sai). Phần lõi so khác và lập kế hoạch vì vậy chắc; lỗi nằm ở chỗ dùng nó (mục 4, 5) và ở mục lục.

Tóm tắt: 2 cao, 6 trung bình, 9 thấp.

---

## CAO

### C1. `word_review_revision` với `all` + `author`: khớp theo "chứa chuỗi", cảnh báo nói "của <tên>" → duyệt nhầm sửa đổi của người khác, không hoàn tác được
- `src/office/wordTools.ts:1557-1569` (lọc `String(t.author).toLowerCase().includes(want)`), cảnh báo ở `:1733-1735`.
- Ví dụ: văn bản có sửa đổi của "Nguyễn Văn An", "Hoàng Anh", "An Phát Law". Người dùng bảo "từ chối hết sửa đổi của An" → AI gọi `{all:true, author:"An"}`. Hộp xác nhận ghi "Từ chối TẤT CẢ sửa đổi đang theo dõi của An", nhưng công cụ từ chối cả của "Hoàng Anh" và "Nguyễn Văn An" (và "An Phát Law"). Thao tác lại bị loại khỏi Hoàn tác (`diffable=false`), nên mất vĩnh viễn sửa đổi của đối tác. Trái quy tắc "không làm hỏng sửa đổi chờ duyệt của người khác".
- Sửa: với duyệt hàng loạt, so khớp **chính xác** tên (không phân biệt hoa thường) hoặc yêu cầu `author` trùng đúng một tên trong `authors` của `word_list_revisions`; nếu chuỗi khớp nhiều tên thì trả danh sách tên và từ chối. Hộp xác nhận ghi rõ số sửa đổi và danh sách tên thật sẽ bị ảnh hưởng (tính trước trong hàm RISK như `word_delete_paragraphs` đã làm).

### C2. `buildOutline`: câu thường mở đầu bằng "Điều N …", "Phụ lục N …", "Phần N …", "Mục N …", "Section N …" bị dựng thành điều khoản / chương → cắt cụt điều đứng trước, đổi cha của các điều sau
- `src/core/outline.ts:80-83` (chương/phụ lục/phần, chỉ chặn `s.length <= 200`), `:85-88` (mục), `:90-97` (điều: không chặn độ dài, không đòi dấu phân cách).
- Đã chạy thử:
  - `["Điều 1. Phạm vi", "Điều 3 của Hợp đồng này quy định rõ quyền của các bên.", "Phụ lục 01 là bộ phận không tách rời của Hợp đồng này.", "Điều 2. Giá"]` cho ra `Điều 1 (end=0)`, `Điều 3 (article, L0)`, `Phụ lục 01 (chapter, L0)`, `Điều 2 (L1, con của Phụ lục 01)`.
  - Hậu quả: `word_read_clause "Điều 1"` chỉ đọc 1 đoạn (đoạn thân bị coi là điều mới); `word_read_clause "Điều 3"` trả điều không có; điều 2 mang cấp 1 và mọi điều sau thành con của "Phụ lục 01"; `word_insert_clause` neo vào `item.end` sai (chèn vào giữa điều khác); `word_find` ghi sai điều chứa kết quả. Câu dẫn chiếu kiểu này là loại câu thường gặp nhất trong hợp đồng ("Phụ lục 01 là bộ phận…", "Điều 12 Luật Nhà ở 2023 quy định…", "Section 3 of this Agreement…").
  - Test `không nhận nhầm …` (`test/unit/dot1a.test.ts:129-140`) chỉ thử số tiền, ngày, năm; không có ca dẫn chiếu.
- Sửa: với số gõ tay (không phải nhãn Word), chỉ nhận "Điều N" khi (a) ngay sau số là dấu `.` `:` `–` `-` hoặc hết dòng, hoặc là khoảng trắng + chữ HOA; và (b) phần còn lại ngắn (≤ ~150 ký tự) và không kết thúc bằng dấu chấm câu của câu dài; hoặc đoạn có kiểu tiêu đề / đậm. Áp dụng cùng quy tắc cho chương, mục, phụ lục (bỏ nhánh chỉ dựa `s.length <= 200`). Thêm test dẫn chiếu.

---

## TRUNG BÌNH

### T1. `buildOutline`: khoản "n." chỉ nhận anh em nếu có khoản `n-1`; khoản đứng trước bị xóa có theo dõi hoặc đánh nhảy số thì khoản sau thành con của khoản trước
- `src/core/outline.ts:227-235` (`it.n === n - 1`), kèm `:169` (`row.deleted` bị bỏ qua).
- Đã chạy thử: `Điều 5 / 1. a / [2. đã xóa có theo dõi] / 3. c` cho `5.1` rồi `5.1.3` (L2). Với danh sách Word tự đánh nhưng khoản nằm ở danh sách riêng, cấp 0 (`listLevel 0`, `listId` khác điều): `Điều 5 / 1. / [2. xóa] / 3.` cũng ra `5.1.3`. Tình huống bản đối tác xóa một khoản là đúng kịch bản dự án nhắm tới. Hệ quả: khóa sai (`khoản 3 Điều 5` không tìm thấy, `5.1.3` là khoản không có), `end` của khoản 1 nuốt khoản 3.
- Sửa: coi `single` trên đường đi có cùng danh sách, cùng cấp và `n` nhỏ hơn `n` hiện tại là anh em (không đòi đúng `n-1`); chỉ lồng vào khoản trước khi `n === 1` hoặc có dấu hiệu cấp sâu hơn (`listLevel`). Thêm test khoản bị xóa/nhảy số.

### T2. `restoreText` (Hoàn tác trường hợp "nhập vào sửa đổi có từ trước") lập kế hoạch không có `spans` → ánh xạ tham lam có thể trỏ vào chữ đã xóa và sửa nhầm chữ
- `src/office/wordTools.ts:711` (`planParagraphEdit(p.text, C, prev)`, không truyền spans); đường dò tham lam ở `src/core/worddiff.ts:124-131`. Ở `word_set_paragraph_text` spans chỉ có khi có `WordApiDesktop 1.4` (`wordTools.ts:1005`); không có thì cũng rơi vào đường tham lam.
- Đã chạy thử: `M = "là 12 tháng18 tháng"` (khúc "12 tháng" của đối tác đã xóa), `C = "là 18 tháng"`, câu mới `"là 20 tháng"` → kế hoạch thay chữ "1" thứ nhất (thuộc khúc ĐÃ XÓA, `at:3`) bằng "20" và xóa chữ "8" sống; kết quả sống là `"là 201 tháng"`. Với `set_paragraph_text` có kiểm lại câu sau khi ghi (`:1048`) thì báo `ok:false` sau khi đã ghi sửa đổi; với `restoreText` không có kiểm lại nên Hoàn tác âm thầm để lại chữ sai.
- Sửa: lấy spans trong `restoreText` giống `word_set_paragraph_text` (tách thành hàm dùng chung); khi M ≠ C mà không có spans thì **từ chối** (trả chuỗi lỗi) thay vì dò tham lam, hoặc dò tham lam chỉ khi mọi vị trí gán là duy nhất (không có ký tự chết trùng ký tự sống). Thêm kiểm lại sau `applyPlan` trong `restoreText` và báo lỗi nếu khác.

### T3. Hoàn tác chọn sửa đổi theo khóa "loại|tác giả|phút|chữ", không theo vị trí: sửa đổi có từ trước trùng khóa với sửa đổi của AI bị từ chối thay
- `src/core/worddiff.ts:250-255` (khóa), `:354-367` (`pickByKeys` luôn lấy mục đứng SAU), dùng ở `wordTools.ts:556`.
- Add-in ghi sửa đổi dưới tên chính người dùng, nên tác giả không phân biệt được sửa đổi của AI với sửa đổi người dùng tự làm trước đó. Ví dụ: 10:05 người dùng xóa có theo dõi "30" ở khoản 3; vẫn phút 10:05 nhờ AI đổi "30" → "45" ở khoản 1 (đứng TRƯỚC trong văn bản). Khóa `Deleted|tôi|10:05|2|30` trùng; `pickByKeys` chọn cái đứng sau = sửa đổi của người dùng. Hoàn tác từ chối nhầm sửa đổi của người dùng và để nguyên sửa đổi của AI. Test `:410-415` chỉ khẳng định trường hợp AI đứng sau.
- Sửa: `addedKeys` trả luôn **chỉ số trong `after`**; `rejectKeys` dùng các chỉ số đó nếu danh sách hiện tại ở các chỉ số ấy vẫn đúng khóa (kiểm lại bằng khóa), chỉ khi lệch mới rơi về tìm theo khóa. Hoặc thêm vào khóa mã đoạn (`getRange().paragraphs.getFirst().uniqueLocalId`).

### T4. Công cụ sửa/chèn/định dạng theo chỉ số đoạn mà `expect` không bắt buộc → vẫn sửa nhầm đoạn khi chỉ số đã dời
- `src/office/wordTools.ts`: `word_replace_text` (`:110-111`, expect tùy chọn), `word_insert_paragraphs` (`:151`), `word_format_text` (`:207`), `word_add_comment` (`:232`), `word_insert_table` (`:248`), `word_go_to` (`:355`, không có expect). `locateParagraph` (`worddiff.ts:391-398`) khi không có `expect` dùng thẳng chỉ số.
- Quy tắc dự án "không sửa nhầm đoạn"; test `dot1a.test.ts:441-447` chỉ ép `expect` bắt buộc cho 4 công cụ. Ví dụ: AI đọc đoạn 40, chèn hai đoạn phía trên, rồi `word_insert_paragraphs {position:"after_paragraph", paragraph_index:40}` không expect → chèn sau đoạn 40 cũ (nay là đoạn khác); tương tự thêm nhận xét, định dạng cả đoạn, chèn bảng.
- Sửa: khi có `paragraph_index` mà thiếu `expect` thì từ chối kèm hướng dẫn (hoặc bắt buộc `expect` trong schema khi `paragraph_index` có mặt, kiểm trong `resolvePara`). Mở rộng test để bắt mọi công cụ có `paragraph_index`.

### T5. `word_review_revision` theo chỉ số: `texts` (đối chiếu) tùy chọn; và liệt kê mọi chỉ số vượt qua cảnh báo "tất cả"
- `src/office/wordTools.ts:1577-1590` (thiếu `texts[k]` thì `exp` rỗng → không kiểm), `:1733-1735` (RISK chỉ xét `a.all`).
- Ví dụ: AI gọi `{action:"accept", indexes:[0,1,2,…,N-1]}` mà không có `texts` → chấp nhận toàn bộ sửa đổi không hỏi, không đối chiếu, không hoàn tác được. Chỉ số cũng lệch dễ dàng vì "mỗi lần AI sửa lại sinh thêm sửa đổi".
- Sửa: bắt buộc `texts` cùng độ dài `indexes` (từ chối nếu thiếu); RISK hỏi khi số chỉ số vượt một ngưỡng nhỏ hoặc bằng/chiếm gần hết `total`.

### T6. `word_insert_clause` (cùng cấp / lên cấp): chèn SAU toàn bộ phạm vi `end` của điều cuối, tức sau cả khối chữ ký không nằm trong bảng
- `src/office/wordTools.ts:1170-1172` (chỉ lùi qua đoạn trong bảng, đoạn đã xóa, đoạn rỗng); `outline.ts:277-286` (điều cuối có `end` = đoạn cuối văn bản; test `dot1a.test.ts:52` khẳng định "khoản cuối kéo tới hết văn bản").
- Ví dụ: Điều 10 là điều cuối, sau đó là `"ĐẠI DIỆN BÊN A"`, `"(Ký, ghi rõ họ tên)"`, `"ĐẠI DIỆN BÊN B"` dạng đoạn thường (không bảng). Chèn điều khoản cùng cấp sau Điều 10 đặt điều mới **sau chữ ký**; cũng đúng với bất kỳ đoạn thường nào đứng sau điều cuối của một chương/phụ lục. Đồng thời đoạn mới đặt `style` của điều neo nhưng nhận danh sách của đoạn cuối (chưa chắc cùng `listId` với điều neo nếu khoản con nằm ở danh sách khác: mã chỉ sửa đoạn nào "không là list item" ở `:1192-1196`, còn đoạn đã là list item của danh sách khác được giữ nguyên danh sách ấy, chỉ đổi cấp ở `:1207-1209`).
- Sửa: với chèn cùng cấp, đặt vị trí sau đoạn con cuối cùng **thực sự thuộc cây điều khoản** (đoạn cuối cùng có mục lục là hậu duệ, cộng các đoạn thường liền ngay sau nó chưa tới mốc không thuộc điều), hoặc dừng ở đoạn đầu tiên giống chữ ký (đoạn trong bảng, dòng trống liên tiếp, kiểu khác). Đọc lại số đã gán và kiểm `listId` của đoạn mới = `anchorRow.listId`; khác thì `attachToList` lại.

---

## THẤP

### L1. `buildOutline`: số thập phân đầu đoạn thành khoản có dấu chấm
- `outline.ts:99-103`. `"1.5 triệu đồng/tháng"`, `"3.5 tỷ đồng"` ở đầu đoạn → `dotted [1,5]` (đã chạy thử: `1.5|dotted|L1`). Chỉ loại `%`, số có nhóm 3 chữ số, 4 chữ số. Sửa: với số gõ tay, đòi nhóm cuối khác đơn vị tiền/độ đo thường gặp, hoặc chỉ nhận khi phần sau bắt đầu bằng chữ hoa / ngắn / có kiểu đậm; hoặc yêu cầu phải có khoản cha cùng tiền tố đang mở (`5.` rồi `5.3`).

### L2. `buildOutline`: danh sách chữ hoa gõ tay C. D. I. L. M. V. X. thành "chương"
- `outline.ts:110-113`. `"C. Công ty X"`, `"D. Công ty Y"` → `chapter` cấp 0 (đã chạy thử), cắt cụt điều đang mở. Sửa: số La Mã hoa gõ tay chỉ nhận khi đã có mục La Mã trước đó (I trước II) hoặc khi chữ cái không là chữ đơn (loại `[CDILMVX]` đơn lẻ trừ `I`, và đòi `I` tiếp theo là `II`).

### L3. `word_list_comments`: nếu nhận xét đầu tiên quá ngân sách (nhiều trả lời) thì không tiến được
- `wordTools.ts:1409-1416`: `takeWithinBudget` không có `shrink` nên khi mục đầu vượt ngân sách `picked` rỗng, `truncated:true`, `next_start = from` (không tiến) → AI gọi lại mãi cùng trang. Sửa: truyền `shrink` cắt `replies` (như `word_list_revisions`), hoặc luôn lấy ít nhất một mục.

### L4. `word_read_paragraphs` với `end` lớn nạp sửa đổi cho cả cửa sổ chứ không chỉ phần trả về
- `wordTools.ts:883-900`: `win = paras.items.slice(start,end)` nạp `getTrackedChanges()` cho từng đoạn trong cửa sổ trong một lần sync; model đặt `end` = tổng số đoạn trên văn bản vài nghìn đoạn có sửa đổi → hàng nghìn lời gọi, trong khi `packParagraphs` chỉ trả vài chục đoạn. Sửa: giới hạn cửa sổ nạp (ví dụ 300 đoạn từ `start`, bằng `end` mặc định).

### L5. `word_replace_text`: cụm thay trùng hệt cụm cũ vẫn sinh sửa đổi xóa/thêm vô nghĩa
- `wordTools.ts:972-977`: khi `replace === t.text` kế hoạch có `hunks = 0`, điều kiện `plan.hunks >= 1` sai nên rơi vào `t.insertText(replace,"Replace")`: ghi xóa + thêm cùng chữ rồi báo "Đã thay". Sửa: nếu `hunks === 0` thì bỏ qua mục tiêu đó (và không tính vào `replaced`).

### L6. Chuỗi tìm có `^` bị nhân đôi khi gửi cho Word nhưng giới hạn 255 chỉ tính trước khi nhân đôi
- `util.ts:146-148` (`^` → `^^`), `worddiff.ts:176` (MAX_SEARCH = 255 tính trên chuỗi gốc), `wordTools.ts:909`, `:656`. Chuỗi 200 ký tự có 60 dấu `^` thành 260 ký tự → Word ném lỗi ở `ctx.sync()` (lỗi xảy ra trước khi ghi nên an toàn, nhưng thông báo khó hiểu). Sửa: tính độ dài sau `escapeWordSearch`.

### L7. Hoàn tác âm thầm thiếu/không đủ
- `wordTools.ts:1671` — nếu `changeSnaps()` đầu lượt lỗi (`before = null`) thì sau lượt không có mục Hoàn tác nào (kể cả khi AI đã ghi) mà không báo gì cho người dùng. Sửa: nếu `before` là null ở công cụ ghi, ghi một mục Hoàn tác báo "không hoàn tác được lượt này" hoặc báo trong kết quả.
- `wordTools.ts:546-581` — `rejectKeys` trả số đã từ chối nhưng bên gọi (`:1683`) bỏ qua: sửa đổi AI đã bị người dùng chấp nhận hoặc từ chối trước đó thì Hoàn tác "thành công" mà không làm gì. Sửa: so với `d.added.length` và báo phần chưa hoàn tác được.
- `wordTools.ts:561-581` — nếu lệnh từ chối cả loạt lỗi sau khi đã từ chối được một phần, nhánh dự phòng dùng lại toàn bộ `keys`, có thể chọn nhầm một sửa đổi trùng khóa có từ trước (liên hệ T3). Sửa: tính lại danh sách cần từ chối từ trạng thái hiện tại so với `before`.

### L8. Công cụ chạy tiếp ngầm sau hết hạn chờ nên có thể đan xen với công cụ kế tiếp
- `loop.ts:197` (`withTimeout` 120 giây, "không hủy được thao tác Office đang chạy") nhưng `execute` (`wordTools.ts:1667-1693`) chụp hai lần danh sách sửa đổi quanh `run`. Công cụ bị hết hạn vẫn chạy ngầm trong khi vòng lặp chạy công cụ kế tiếp; lần chụp của lượt này bắt luôn sửa đổi của công cụ kia → mục Hoàn tác nhập lẫn hoặc thiếu. Sửa: trong `execute` giữ khóa tuần tự (một hàng đợi ghi): công cụ kế không bắt đầu khi công cụ cũ chưa xong (kể cả khi vòng lặp đã bỏ chờ).

### L9. `word_read_tables` đọc `t.values` (có chữ đã xóa có theo dõi), không theo nguyên tắc "bản hiện tại" của các công cụ đọc khác — chưa chắc
- `wordTools.ts:1599-1627`. `paragraph.text` đã được kiểm chứng gồm cả chữ đã xóa; `table.values` rất có thể cũng vậy, nên AI đọc ô như "24 tháng12 tháng". Cùng giá trị cũ đi vào `planTableFill` (`util.ts:132-135`) khiến cảnh báo ghi đè đếm nhầm. Đề nghị kiểm thử trong Word thật; nếu đúng thì đọc ô bằng `getText()`/`getReviewedText("Current")` như đoạn.

---

## CHƯA CHẮC (cần kiểm bằng Word thật, không tính vào số phát hiện)

- `word_format_text` (`wordTools.ts:1269-1307`) và `word_change_level` (`:1238-1257`) có được ghi dưới dạng sửa đổi theo dõi không? Định dạng chỉ được theo dõi khi tùy chọn "Track formatting" của Word bật, đổi cấp danh sách bằng `listItem.level` có thể không sinh `TrackedChange`. Nếu không sinh, sửa đổi trái quy tắc "mọi sửa đổi ghi bằng Track Changes" và `execute` không ghi mục Hoàn tác nào (`d.added` rỗng, `u.extra` rỗng). Selftest "Đổi cấp điều khoản" (`selftestWord.ts:407-412`) và "Định dạng chữ" (`:414-432`) không kiểm sửa đổi theo dõi hay Hoàn tác.
- Chuỗi tìm toàn khoảng trắng hoặc có khoảng trắng ở đầu/cuối làm neo (`planParagraphEdit` tạo neo như `"toán "`, `" "`; `worddiff.ts:229-233`, và test `dot1a.test.ts:336`). Selftest chỉ có các chỗ thay có chữ; không có ca chèn thuần neo bằng chuỗi kết thúc bằng khoảng trắng. Nếu `Range.search` của Word bỏ khoảng trắng đầu/cuối thì chữ chèn lệch một ký tự và chỉ bị phát hiện ở bước kiểm lại sau khi đã ghi.
- `removeParas` (`wordTools.ts:718-736`) xóa đoạn theo `uniqueLocalId` khi chế độ theo dõi Tắt; nếu mã đoạn bị Word cấp lại sau khi người dùng gõ vào/ghép đoạn thì có thể xóa nhầm đoạn khác không để lại dấu vết. Chỉ chạy ở nhánh "nhập vào sửa đổi có từ trước". Nên kiểm tra thêm nội dung đoạn (đầu đoạn khớp với chữ AI đã chèn) trước khi xóa.
- `snapshot` (`wordTools.ts:1707`) ghi nhãn số vào trước chữ đoạn để ứng dụng khác đọc; nếu AI chép nguyên dòng đó làm `expect` thì `matchesExpect` sẽ từ chối (nhãn không nằm trong chữ đoạn). Chỉ ảnh hưởng nếu snapshot được đưa cho phiên Word.
- Thay cả đoạn bằng `insertText(raw, "Replace")` (`wordTools.ts:1027`, `:1032`) có thể làm mất đối tượng nội tuyến (ảnh, trường, tham chiếu chéo, ghi chú cuối trang) trong đoạn; chỉ là xóa/thêm có theo dõi nên người dùng từ chối được, nhưng cảnh báo không nêu.

---

## Kiểm thử máy nào quá lỏng hoặc thiếu (mục 5)

1. `dot1a.test.ts:129-140` (không nhận nhầm): không có ca dẫn chiếu ("Điều 3 của…", "Phụ lục 01 là…"), số thập phân đầu đoạn, C./D. gõ tay → để lọt C2, L1, L2.
2. `dot1a.test.ts:301-330` (kế hoạch với chữ đã xóa): cả hai test gọi `planParagraphEdit(M, C, N)` không có `spans`, đúng nhánh dò tham lam, và chọn dữ liệu mà chữ đã xóa không trùng chữ sống ("24" so với "12") nên tham lam may mắn đúng. Nhánh spans chỉ được thử gián tiếp qua `liveMap`. Test `:358-363` còn xác nhận `liveMap("ab ab","ab")` = `[0,1]` — chính là phép gán sai nếu khúc chết là "ab" đầu (T2).
3. `dot1a.test.ts:402-438` (Hoàn tác): chỉ thử các hàm thuần (`addedKeys`, `pickByKeys`, `diffChanges`), không có ca khóa trùng giữa sửa đổi có từ trước và của AI mà AI đứng trước (T3). `rejectKeys`, `restoreText`, `removeParas`, `execute` không có kiểm thử máy; chỉ có Tự kiểm tra trong Word.
4. `dot1a.test.ts:441-447`: chỉ ép `expect` bắt buộc cho 4 công cụ; không bắt công cụ nào khác có `paragraph_index` mà thiếu `expect` (T4).
5. Không có kiểm thử cho khoản bị xóa/nhảy số (T1) và không có kiểm thử ngẫu nhiên cho `applyHunks(old, diffWords(old,new)) === new` (đã thử 200.000 ca, đúng; nên đưa một bản gọn vào `npm test` làm rào chắn).
6. Tự kiểm tra Word không có ca: duyệt theo tên tác giả khớp một phần (C1), nhập chữ chèn thuần neo bằng chuỗi có khoảng trắng đầu/cuối, định dạng/đổi cấp sinh sửa đổi theo dõi và có Hoàn tác.

---

## Điều làm tốt (để giữ)

- Tách ba thứ rõ ràng: chuỗi đánh dấu M (để tìm), bản hiện tại C (để so), spans (vị trí chữ đã xóa); lõi `diffWords`/`planParagraphEdit` đúng trên 200.000 ca ngẫu nhiên.
- `locator` chỉ nhận chuỗi mà lần xuất hiện bắt đầu đúng vị trí theo cách đếm không chồng của Word, kèm `total` để đối chiếu số kết quả Word trả; lệch thì từ chối trước khi ghi.
- `tracked()` luôn trả lại chế độ theo dõi trong `finally`; `execute` ghi mục Hoàn tác cả khi công cụ ném lỗi giữa chừng.
- Ngân sách kết quả tính theo độ dài JSON thật và có `next_start`.
