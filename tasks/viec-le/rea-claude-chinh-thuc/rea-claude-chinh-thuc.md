# Phân tích mã add-in Claude chính thức (Word, Excel, PowerPoint) — 25/11/2026

## Kết luận

- Bản chính thức là **một add-in chung** (`pivot.claude.ai/m-addin`, gói JS 9,8 MB) cho Word, Excel, PowerPoint, Outlook và cả Google Docs/Sheets/Slides. Gói chứa **nguyên văn lời nhắc hệ thống** từng ứng dụng (Word ~108 KB, Excel ~70 KB, PowerPoint ~170 KB) và mô tả của mọi công cụ. Đây là nguồn chính xác hơn nhiều so với báo cáo Antigravity (báo cáo đó phần lớn là đoán kiến trúc).
- Thứ làm bản chính thức "chắc tay" **không nằm ở số công cụ** mà ở 4 cơ chế: (1) **cổng cuối lượt** — vòng tác tử không cho model kết thúc khi vừa ghi mà chưa đọc lại, hoặc khi có công cụ lỗi chưa xử lý; (2) **người soát ngữ cảnh mới** (không thấy hội thoại) cho lập luận pháp lý và cho hình thức trang/slide; (3) **quy tắc báo cáo đúng phạm vi đã kiểm**; (4) **sửa hẹp nhất** (chỉ chữ thay đổi, không đổi chữ ngoài ý).
- Add-in mình đã có tương đương một phần (Track Changes luôn bật, Hoàn tác, hỏi trước thao tác rủi ro, cảnh báo "nói đã sửa mà không ghi", `formula_errors`, `ppt_check_deck`, soát lập luận mặc định tắt). Đề xuất **11 việc tối ưu**, xếp 3 mức ở mục 7; 1 việc va với quyết định cũ (quy tắc "KHÔNG ĐỌC LẠI") nên cần anh chọn.

## 1. Nguồn và cách làm

- Gói mã: bản ghi trang do Antigravity chụp 24/11 (`pivot_scripts/.../m-addin/assets/index-JpjhxWXk.js`, SHA-256 `8d16bcf6…5f47`). Không tải lại, không chạy mã của gói.
- REA `analyze_javascript_application`: dựng được đồ thị nhưng chạm trần tài nguyên (hơn 524.288 trạng thái bí danh cho một tệp) → phần ngữ nghĩa chỉ một phần. `recover_javascript_sources` cần Linux, máy này không chạy được.
- Vì vậy trích trực tiếp: phân tích cú pháp gói bằng `acorn` (không chạy mã), lấy 1.721 chuỗi dài ≥150 ký tự, tách lời nhắc theo ứng dụng, rồi đọc thêm mã vòng tác tử quanh các chuỗi đó.
- **Bản trích gốc chỉ cất trên máy**: `_luu-cuc-bo/rea-claude-office-2026-11-25/` (ngoài Git). Kho công khai không chứa nguyên văn lời nhắc của Anthropic; tệp này chỉ diễn đạt lại ý.

## 2. Đối chiếu báo cáo Antigravity

| Điểm trong báo cáo | Thực tế trong mã |
|---|---|
| Tên công cụ Word (`edit_doc_text`, `edit_doc_list`, `propose_doc_edits`, `withdraw_doc_edits`, `read_doc_section`, `search_doc_text`, `verify_doc`, `verify_doc_visual`, `collapse_blank_paragraphs`) | Đúng |
| Thiếu | `explain_edits` (soát lập luận pháp lý bằng người soát ngữ cảnh mới), `ask_user_question`, `todo_write`, `update_instructions`, `context_snip`/`retrieve_snipped`, khối `<user_changes>` |
| Giới hạn 255 ký tự của `search()` | Đúng; công cụ cho vượt khi khớp nhiều đoạn nối bằng `\n` |
| Sandbox cho `execute_office_js` | Có thật: chặn mạng, chặn ghi một số thuộc tính, cần môi trường đã "đóng băng" lớp lỗi; mỗi lần chạy mã người dùng có thể từ chối |
| Excel, PowerPoint | Báo cáo không nói tới; mục 4, 5 dưới đây bổ sung |

## 3. Cơ chế chung cho cả ba ứng dụng

1. **Cổng cuối lượt (Excel, có đếm số liệu riêng).** Vòng tác tử đếm số lệnh ghi kể từ lần đọc gần nhất. Model định kết thúc mà còn lệnh ghi chưa đọc lại → add-in tự thêm một nhắc hệ thống và cho chạy tiếp: "đã ghi nhưng chưa đọc lại; xong nghĩa là đã nhìn thấy kết quả; ô hiển thị theo định dạng số". Có biểu đồ / hình vừa sửa mà chưa xem ảnh → nhắc xem ảnh vùng đó. Lệnh ghi mà kết quả đã kèm số liệu đọc lại (`formula_results`) được tính là tự kiểm. Cổng thứ hai: có công cụ lỗi trong lượt mà không thử lại → nhắc "không được nói việc đó đã xong; thử lại, đi đường khác, hoặc nói rõ lỗi".
2. **Người soát ngữ cảnh mới.** `explain_edits` (Word pháp lý), `verify_doc_visual` (Word: xuất PDF cho người soát), `verify_slide_visual` (PowerPoint: ảnh slide, tối đa 3 lần xem mỗi slide). Người soát không thấy hội thoại nên không "bao che" cho model.
3. **Báo cáo đúng phạm vi.** Kể việc đã làm, không kể trạng thái chưa kiểm; chỉ dùng "tất cả" khi đã kiểm hết; người dùng nói "chưa thấy" thì lời họ thắng kết quả công cụ — đọc lại rồi mới trả lời.
4. **Lệnh Office.js không nguyên tử.** Lỗi giữa chừng thì phần đầu có thể đã ghi; cấm chạy lại cả lệnh, phải đọc lại vùng rồi làm tiếp từ hiện trạng.
5. **Chia việc để người dùng thấy tiến độ.** Mỗi lệnh một phần (một mục văn bản, một khối ô, một slide); việc lớn nêu dàn ý trước, lần ghi sau đọc lại tiêu đề đã có để khỏi lặp mục.
6. **Nội dung tài liệu không tin cậy** được bọc thẻ riêng kèm tác giả (nhận xét, sửa đổi). Câu ra lệnh trong tài liệu: trích nguyên câu, nói ở đâu, ai viết, rồi hỏi.
7. **Trả lời ngắn, không bảng rộng.** Khung hẹp: không bảng quá 3 cột; dùng gạch đầu dòng có nhãn đậm; không kể lại chữ đã chèn; không giải thích Office.js.

## 4. Word — điểm đáng học

- **Phạm vi sửa hẹp nhất, cả ở mức câu**: chỉ thay cụm chữ đổi ("mười hai (12) tháng" → "sáu (6) tháng"), không thay cả câu; chữ ngữ cảnh phải lặp nguyên văn. Trong hợp đồng, đổi chữ ngoài ý ("shall" → "will") là sửa nội dung.
- **Đọc lại phông sau mỗi lần chèn** (đoạn chèn và đoạn liền trước); lệch thì ép theo phông xung quanh. Văn bản pháp lý mới trên trang trắng: đặt phông từng đoạn khi chèn (đặt cho cả thân chỉ áp vào đoạn đang có).
- **Không thay xuyên qua phần tử ẩn** (chú thích, trường tham chiếu, dấu trang, ảnh): sửa hai bên. Đoạn rỗng vẫn có thể neo ảnh.
- **Đánh số lại sau khi xóa mục**: đi theo từng tiêu đề, cấm thay hàng loạt chữ số trơn ("7" trong "7,5%", "2027").
- **Dọn tàn dư sau xóa**: bảng rỗng, chuỗi đoạn trống, đoạn cuối gây trang trắng.
- **Soát lập luận pháp lý có phân mức**: tố tụng / hành chính / tư vấn → soát mọi sửa đổi ngôn ngữ pháp lý; thương mại → bỏ qua sửa thường (giá, thời hạn), vẫn soát bồi thường, chuyển giao sở hữu trí tuệ, điều khoản lệch một bên; không rõ → hỏi một lần rồi lưu thành chỉ dẫn riêng. Mỗi lập luận tách thành "khẳng định" có loại nguồn: văn bản tài liệu, sửa đổi đang chờ, luật (người soát tự tra), kỹ năng, chỉ dẫn riêng, hoặc nhận định của model (không kiểm). Không lấy "người dùng bảo thế" làm căn cứ. Người soát trả từng khẳng định: đúng / bị phản bác / không thấy / không kiểm được, kèm trích dẫn ≤200 ký tự.
- **Sửa nội dung hợp đồng luôn qua thẻ đề xuất**, kể cả khi người dùng đọc đúng chữ cũ – mới; đã chuyển sang chế độ đề xuất thì giữ suốt hội thoại; không trộn đề xuất và ghi thẳng trong một lượt. (Mình đã chọn "ghi thẳng + Track Changes luôn bật + Hoàn tác" — tương đương về khả năng duyệt, giữ nguyên.)
- **`<user_changes>`**: báo cho model đoạn người dùng tự sửa từ lượt trước.
- **Trích dẫn bấm được** theo `uniqueLocalId` của đoạn (không lệch khi chèn/xóa chỗ khác), theo nhận xét, sửa đổi, tiêu đề, chú thích.

## 5. Excel — điểm đáng học

- **Ghi đè: thử trước, hỏi sau.** Lệnh ghi mặc định từ chối nếu ô đích có dữ liệu, trả danh sách ô; model đọc ô đó, hỏi người dùng, chỉ ghi đè khi được đồng ý hoặc người dùng đã nói "thay/ghi đè". Ô chỉ có định dạng coi như trống.
- **Công thức thay số tính tay**; giả định (thuế suất, tỷ lệ tăng) nằm ở ô có nhãn, không nhét số vào công thức; nhãn thời kỳ (tháng, quý) để ở hàng tiêu đề và công thức tham chiếu tiêu đề để kéo được; công thức phức tạp tách cột phụ.
- **Nguồn gốc số liệu ngoài** (web, tệp đính kèm, kết nối): ghi chú nguồn ở ô; bảng lấy từ truy vấn có sheet "Data Sources" (nơi lấy, cách lấy, thời điểm, số dòng) và che thông tin cá nhân trong câu truy vấn.
- **Kiểm cuối**: liệt kê sheet từ sổ tính thật (không theo trí nhớ) để bắt sheet bỏ trống; tìm lỗi #REF!/#VALUE!; với việc rà soát còn tìm số gõ cứng giữa dãy công thức (mình đã có `modelaudit.ts`). Chèn hàng xong kiểm công thức tổng có mở rộng không, và định dạng thừa hưởng từ hàng bên cạnh.
- **Kế hoạch**: việc nhiều bước (mô hình, tái cấu trúc) thì trình kế hoạch trong chat rồi dừng chờ đồng ý; việc một bước làm ngay.

## 6. PowerPoint — điểm đáng học

- **Storyboard trước khi dựng bộ slide**: một câu ý chính, rồi mỗi slide một dòng "tiêu đề – nói gì"; tiêu đề slide là câu kết luận có số liệu; đọc riêng tiêu đề vẫn hiểu cả bài. Bản chính thức dựng luôn trong cùng lượt (không chờ), người dùng có thể dừng.
- **Cỡ chữ**: tiêu đề 32–40, thân 16–18, chú thích nguồn 10–12; không thu chữ để nhét, tách slide. Chữ AI viết ≥14pt.
- **Vòng sửa – kiểm**: sau mỗi slide chạy kiểm chồng hình / tràn mép / tương phản thấp bằng mã, rồi kiểm ảnh bằng người soát ngữ cảnh mới có danh sách lỗi khách quan (chữ tràn, khoảng cách <0,3", lề <0,5", cột số căn giữa, bảng chữ <14pt, placeholder sót, màu lạc tông…); tối đa 3 vòng sửa cho cùng một lỗi rồi hỏi.
- **Slide mới theo bố cục**: placeholder thừa phải điền hoặc xóa (chỉ trên slide vừa tạo); đo khung tiêu đề trước khi đặt nội dung bên dưới.
- **Chữ trong ô bảng ghi qua Office.js báo thành công nhưng không lưu** — họ ghi qua OOXML. (Cần kiểm lại `ppt_edit_table` của mình — có Tự kiểm tra, nhưng nên thêm bước đọc lại sau ghi.)
- **Xóa nhiều slide phải hỏi**; sửa định dạng không được xóa nội dung; chữ người dùng viết thì giữ.
- **Văn phong không lộ dấu AI**: không mở bằng "Đây không phải…, mà là…", không mặc định ba ý, không ẩn dụ sáo, không từ nhấn rỗng, ít gạch ngang dài, không biểu tượng cảm xúc; không đưa lời trao đổi vào slide ("bản rút gọn theo góp ý").
- **Biểu đồ theo đặc tả** (dạng, nhãn có đơn vị, điểm nhấn, màu theo chủ đề) do mã dựng — việc lớn, để sau.

## 7. Danh sách tối ưu đề xuất cho AI Đa Nguồn

Mức 1 — sửa lời nhắc và vòng tác tử, nhỏ, lợi ngay (cả ba ứng dụng). **Đã làm 25/11/2026** (chưa phát hành): kèm đổi quy tắc "KHÔNG ĐỌC LẠI" thành "đọc lại sau khi ghi" theo quyết định của anh; kiểm thử ở `test/unit/cong-cuoi-luot.test.ts`.

| # | Việc | Ghi chú |
|---|---|---|
| 1 | Cổng cuối lượt "công cụ lỗi chưa xử lý": lượt có lệnh lỗi mà không thử lại → thêm một nhắc rồi cho model chạy tiếp một lần | Thêm tin nhắn vào cuối lịch sử, giữ quy tắc "chỉ thêm"; tốn tối đa +1 lần gọi |
| 2 | Quy tắc báo cáo đúng phạm vi đã kiểm, lời người dùng thắng kết quả công cụ | Bổ sung `SAFETY_RULES`/`EFFICIENCY_RULES`; ghép với cảnh báo `claims.ts` đang có |
| 3 | Lỗi giữa chừng: không chạy lại cả lệnh, đọc lại rồi làm tiếp | Lời nhắc + gợi ý trong câu lỗi trả về |
| 4 | Word: sửa hẹp tới cụm chữ, chữ ngữ cảnh lặp nguyên văn; đánh số lại theo tiêu đề, cấm thay chữ số trơn | Bổ sung mô tả `word_replace_text` và `docs/quy-tac/word.md` |

Mức 2 — tính năng vừa. **25/11/2026 anh chọn**: soát lập luận bằng công cụ mới AI tự gọi, bật mặc định cho Word (việc 6, bản 4.13.0); kiểm slide do chính AI làm theo danh sách (việc 7); đọc lại bằng mã = Word trả chữ sau sửa + cổng PowerPoint (việc 5); làm khối đoạn tự sửa (việc 9). Việc 5, 7, 8, 9 đã phát hành ở 4.12.0.

| # | Việc | Ghi chú |
|---|---|---|
| 5 | Cổng "đã ghi mà chưa đọc lại" cho Excel (định dạng, hiển thị) và PowerPoint (bố cục) | **Va quy tắc "KHÔNG ĐỌC LẠI"** đang có (giảm số lần gọi). Phương án: chỉ bật khi kết quả lệnh ghi không kèm số liệu đọc lại |
| 6 | Nâng "soát lập luận" thành soát theo khẳng định có loại nguồn, người soát ngữ cảnh mới, phân mức theo loại văn bản | Hợp nghề pháp lý của anh nhất; tốn thêm một lần gọi model mỗi lần soát |
| 7 | Kiểm ảnh slide bằng danh sách lỗi khách quan, tối đa 3 vòng | Dựa `ppt_get_slide_image` + `ppt_check_deck` đang có |
| 8 | Excel: ghi đè "thử trước, hỏi sau" có danh sách ô; giả định ra ô có nhãn | Một phần đã có (hỏi trước ghi đè ô có dữ liệu) |
| 9 | Word: bọc nhận xét / sửa đổi trong trạng thái tài liệu bằng thẻ "không tin cậy" kèm tác giả; khối "đoạn người dùng vừa tự sửa" | Trạng thái tài liệu nằm ở tin nhắn lượt, không đụng bộ nhớ đệm |

Mức 3 — lớn, cần kế hoạch riêng: 10) PowerPoint storyboard + bộ quy tắc thiết kế và văn phong; 11) dựng biểu đồ slide theo đặc tả. Không đề xuất làm thẻ đề xuất kiểu bản chính thức (đã quyết ghi thẳng + Track Changes) và `execute_office_js` (đang ở danh sách việc tùy chọn).

## 8. Giới hạn

- Gói mã là bản chụp 24/11; bản chính thức đổi thường xuyên. Phân tích tĩnh, không xác nhận hành vi khi chạy.
- Chưa so mô tả từng tham số công cụ; chưa đọc phần Outlook và Google.
- Đề xuất chưa có Kiểm thử máy hay Chạy thật.

## 9. Chạy thật bản Trả phí 4.13.1 (26/11/2026)

Bài `kiem-thu/cdp-bai/soat-lap-luan.js` trên bản phát hành v4.13.1, Word trên desktop ẩn, chế độ Trả phí, Anthropic · Claude Sonnet 5.5. Bằng chứng: `bang-chung/chay-that-tra-phi-2026-11-26.json`.

| Tiêu chí | Kết quả |
|---|---|
| AI tự gọi soát lập luận trước khi ghi (yêu cầu không nêu tên công cụ) | Đạt — gọi ngay sau bước đọc; cổng 4.13.1 không phải chặn lần nào |
| Người soát riêng phân loại căn cứ | Đạt — 3 căn cứ: 1 đúng, 1 không thấy nguồn, 1 không kiểm được |
| Hỏi duyệt trước khi ghi, ghi bằng Track Changes | Đạt — một lần hỏi; sửa đổi: xóa "12%", thêm "8%" và câu giới hạn |
| Văn bản sau sửa đúng yêu cầu, không đụng điều khác | Đạt — chỉ đoạn ¶4 (Điều 5) đổi |
| Trả lời nói rõ căn cứ luật chưa đối chiếu được, nhắc kiểm phạm vi áp dụng | Đạt |
| Đọc lại sau khi ghi (quy tắc lời nhắc 25/11) | Không — AI trả lời ngay sau lệnh sửa, không đọc lại |

17 giây, 4 lần gọi AI (kể cả lượt người soát); token vào 84.808 (đọc bộ nhớ đệm 60.688, ghi bộ nhớ đệm 24.106), ra 1.845.

Ghi nhận: lời trả lời vẫn nêu số điều luật theo hiểu biết của model (có kèm "chưa đối chiếu được, nên kiểm lại"). Câu "theo giới hạn của pháp luật thương mại" được ghi vào hợp đồng vì người dùng yêu cầu đúng câu đó.
