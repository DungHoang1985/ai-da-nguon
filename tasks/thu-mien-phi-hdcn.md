# Thử model miễn phí: soạn hợp đồng chuyển nhượng QSDĐ kèm nhà xây sẵn (23/11/2026)

**Đề bài** (gửi nguyên văn cho mọi model, mỗi model một hội thoại mới, tài liệu Word trắng):
"Soạn một hợp đồng chuyển nhượng quyền sử dụng đất giữa công ty Anh Quân cho công ty Xây dựng 3 theo quy định mẫu của Luật nhà ở (biết rằng đây là chuyển nhượng quyền sử dụng đất và tài sản là nhà xây sẵn trên đất đã có hạ tầng cơ bản để vào ở)."
Nếu model chưa ghi gì vào Word thì gửi thêm một câu: "Những thông tin chưa có thì bạn tự giả định hợp lý hoặc để trống "…" cho tôi điền sau. Cứ soạn luôn hợp đồng vào tài liệu Word."

**Cách chạy**: add-in v4.11.9 (bản phát hành), chế độ Miễn phí, Word mở trên desktop ẩn và điều khiển qua cổng gỡ lỗi, các lần hỏi quyền ghi đều bấm "Cho phép hết lần này". Nút Web đang tắt (với nguồn miễn phí, add-in không có công cụ tra web khi AI tự thao tác). Bằng chứng gồm tệp .docx, hội thoại và định dạng từng đoạn, nằm ở `tasks/thu-mien-phi-hdcn-bang-chung/`.

## Chuẩn để chấm

- **Căn cứ hiện hành**: Luật Nhà ở 2023 (27/2023/QH15), Luật Đất đai 2024 (31/2024/QH15), Luật Kinh doanh BĐS 2023 (29/2023/QH15), Bộ luật Dân sự 2015, Nghị định 96/2024/NĐ-CP. Các mẫu hợp đồng kinh doanh BĐS nằm ở phụ lục của nghị định này, trong đó có mẫu chuyển nhượng QSDĐ đã có hạ tầng kỹ thuật trong dự án (Phụ lục VII) và nghị định đã được Nghị định 54/2026 sửa một phần. Luật Nhà ở không có mẫu hợp đồng. Luật Nhà ở 2014, Luật Đất đai 2013 và Nghị định 99/2015 đều đã hết hiệu lực.
- **Suy luận mong đợi**: bên bán là công ty, nhà xây sẵn trên đất đã có hạ tầng, nên nhiều khả năng đây là giao dịch kinh doanh BĐS trong dự án. Model nên nêu Luật KDBĐS và mẫu của Nghị định 96, hoặc hỏi xem Anh Quân có phải chủ đầu tư dự án không. Đây là chỗ đề bài cố ý để ngỏ.
- **Thể thức**: quốc hiệu và tiêu ngữ, tên hợp đồng, số, căn cứ in nghiêng, thông tin hai bên, điều khoản, phần ký hai cột.

## Kết quả

| Model | Thời gian · số lần gọi AI | Nội dung | Thể thức | Hỏi lại | Tự chủ suy luận | Tra cứu / căn cứ |
|---|---|---|---|---|---|---|
| **Gemini 3.8 Flash** (Google) | 85 giây · 3 | **7/10**. 10 điều gọn, đủ ý: thanh toán 3 đợt với đợt cuối gắn việc cấp sổ, rủi ro chuyển khi bàn giao, chi phí đăng ký. Lỗi: "QUYẾN", "meter"; tự thêm "TNHH" cho cả hai công ty; ghi "Tòa án cấp huyện" (đã bỏ từ 7/2025) | **8/10**. Đúng bố cục; phần ký "ĐẠI DIỆN BÊN A / GIÁM ĐỐC" đúng mẫu | Không, soạn ngay | 6/10. Bám đề, có ý bảo vệ bên mua, không nhận ra bối cảnh dự án | 3/10. Luật Nhà ở 2014, Đất đai 2013 (hết hiệu lực), không ghi số hiệu, có để "[cần bổ sung]" |
| **Nemotron 3 Ultra** (OpenRouter) | 126 giây · 5 | **7/10**. Dài và đủ nhất (13 điều): tách giá đất và nhà, danh mục hồ sơ bàn giao, cam kết, phạt, bất khả kháng. Lỗi: sót chữ Anh "burden" hai lần, "Bền A", "quyềnđảm bảo", trích "Luật Hợp đồng" (không có luật này), Điều 13 trùng khoản 12.1 | **7/10**. Bố cục đúng; "ĐẠI DIỆN BÊN A" lặp hai dòng ở phần ký, thừa dòng "---" | Không, soạn ngay | 7/10. Tự phát hiện và sửa 2 chữ Hán; chưa nhận ra bối cảnh KDBĐS | **1/10**. Căn cứ toàn văn bản hết hiệu lực và **bịa** "Nghị định 01/2024 sửa Nghị định 99/2015" |
| **ERNIE 5.5 Flash** (OpenRouter) | 8 + 17 giây · 2 + ? | **7/10** (bản trong khung chat). Suy luận đúng hướng nhất: nêu Luật KDBĐS, chủ trương đầu tư, quyết định giao đất, giấy phép xây dựng, biên bản nghiệm thu hạ tầng, bảo hành nhà và hạ tầng. Lỗi chữ "TÀI SẠN" | **0/10**. **Không ghi gì vào Word** qua cả hai lượt | Có. Lượt 1 chỉ đưa đề cương rồi hỏi "Bạn muốn tôi soạn…?" | 8/10 về nội dung. Nhưng **nói sai sự thật**: hai lần bảo "công cụ đang lỗi" mà không hề gọi công cụ | 5/10. Nêu đúng nhóm luật (gồm KDBĐS), không bịa số hiệu, để "…" |
| **Nemotron 3.5 Super VL** (OpenRouter) | 136 giây · 5 | **4/10**. Sơ sài: không có phạt vi phạm, thuế phí, hồ sơ bàn giao. Lẫn chữ lạ "cấmélé", "competencia", "chủthể", "viet" | **4/10**. Phần ký lộn: cột trái ghi Bên A rồi Bên B, cột phải ghi Bên B rồi Bên A | Không | 4/10. Có tự sửa chữ Hàn "체가" nhưng sót chữ Latinh lạ | 4/10. Không bịa, chỉ để "[cần bổ sung căn cứ…]", tức là không có căn cứ nào |
| **Cohere Command A Plus** | 40 giây | **4/10**. 8 điều sơ sài; bắt **bên mua** bảo lãnh ngân hàng; "thời hạn sử dụng đất tính từ ngày bàn giao" (sai bản chất); thiếu bàn giao, phạt, thuế phí | **3/10**. Tự chèn thêm điều nhưng để lại "## Điều 8" dạng markdown, hai Điều 8, Điều 8 đứng trước Điều 7; vẫn báo "Hợp đồng đã sẵn sàng để ký kết" | Không | 4/10. Chủ động bổ sung (ý tốt) nhưng làm hỏng văn bản | 2/10. Ghi đúng số hiệu nhưng là văn bản đã hết hiệu lực (65/2014, 99/2015, 45/2013) |
| **Cohere Command A Reasoning** | 18 + 24 giây | **3/10**. Tự đặt tỷ lệ 30/50/20%, cho đợt 2 trả 50% trước khi bàn giao; tự đặt "Công ty CP Xây dựng 3" | **3/10**. Lượt 2 dán cả bản viết trong chat vào thân hợp đồng nên có **hai phần ký**, đoạn "Lưu ý: cần đính kèm" nằm trong hợp đồng, chỗ trống lẫn "[…]" và "…" | Không hỏi, nhưng lượt 1 chỉ viết trong chat, phải nhắc mới ghi vào Word | 3/10 | 2/10. Luật Nhà ở 65/2014 (hết hiệu lực) |
| Gemma 4 31B | không chạy được | — | — | — | — | — |

Gemma 4 31B được thử 3 lần (2 lần qua OpenRouter, 1 lần qua Google), cả 3 lần nguồn đều hết lượt hoặc quá tải nên add-in tự đổi sang Cohere. Chân khung vẫn ghi "Gemma", nên bản đó không tính cho Gemma; bằng chứng để ở `loai-xoay-vong/`.

## Kết luận

1. **Xếp hạng**: Gemini 3.8 Flash ≈ Nemotron 3 Ultra > ERNIE (suy luận tốt nhất nhưng không ghi được vào Word và nói sai) > Nemotron 3.5 Super > Cohere A Plus > Cohere A Reasoning. Nếu cần một bản nháp dùng được ngay trong Word thì chọn Gemini 3.8 Flash (ít lỗi, nhanh) hoặc Nemotron 3 Ultra (đủ ý hơn nhưng phải soát chữ lạ và xóa phần căn cứ).
2. **Thể thức** của 5 model có ghi vào Word đều đúng khung (quốc hiệu, tên, số, căn cứ nghiêng, ký hai cột). Phần này do công cụ `hop_dong` của add-in lo; lỗi thể thức còn lại đến từ việc model đưa dữ liệu sai vào công cụ.
3. **Hỏi lại**: chỉ ERNIE hỏi (một lần), các model khác soạn ngay. Không model nào hỏi câu đáng hỏi nhất: Anh Quân có phải chủ đầu tư dự án không, vì nếu đúng thì phải dùng mẫu của Nghị định 96/2024.
4. **Tra cứu**: không model nào tra được. Với nguồn miễn phí, add-in không có công cụ tra web khi AI tự thao tác. **Cả 6 model đều không nêu đúng căn cứ hiện hành**; có 4 model ghi văn bản đã hết hiệu lực và 1 model bịa văn bản sửa đổi. Không model nào chỉ ra rằng mẫu hợp đồng không nằm ở Luật Nhà ở mà ở Nghị định 96/2024.

## Lỗi của add-in tìm thấy qua bài thử

**Đã sửa ở 4.11.10** (24/11/2026, cùng các lỗi A8–A21 của buổi văn bản hành chính; chi tiết và kiểm chứng ở `tasks/thu-mien-phi-vbhc.md`). Chạy lại đề này trên bản sửa: căn cứ đúng luật đang có hiệu lực, phần ký không lặp hay lộn bên, model tự sửa từ lạ theo `check_now`.

| # | Lỗi | Thấy ở |
|---|---|---|
| A1 | Sau khi tự đổi nguồn, chân khung vẫn ghi model đã chọn chứ không ghi model thực sự trả lời | Gemma → Cohere |
| A2 | Phần ký: model đặt chức danh trùng tiêu đề thì in "ĐẠI DIỆN BÊN A" hai dòng; model đặt lộn thì in lộn A/B giữa hai cột. Công cụ chưa bỏ dòng trùng, chưa kiểm cột | Nemotron 3 Ultra, Nemotron 3.5 Super |
| A3 | Kiểm chữ lạ (`foreign_text`) chỉ bắt chữ Hán, Hàn, Ả Rập; từ Latinh không phải tiếng Việt lọt qua ("burden", "competencia", "cấmélé", "meter") | Nemotron, Gemini |
| A4 | Đoạn chèn có tiêu đề markdown "## " được ghi nguyên chữ, không đổi thành tiêu đề điều | Cohere A Plus |
| A5 | Body có phần ký và "Lưu ý" nằm giữa hoặc cuối (bản dán từ chat) mà không bị bỏ, nên văn bản có hai phần ký | Cohere A Reasoning |
| A6 | Model nói "công cụ đang lỗi" mà không gọi công cụ, hoặc nói "đã sẵn sàng ký" khi văn bản còn hỏng: chưa có cảnh báo (`claims.ts` mới bắt câu "đã sửa") | ERNIE, Cohere A Plus |
| A7 | Kỹ năng `soan-hop-dong` chưa có danh mục căn cứ hiện hành cho hợp đồng BĐS, nên model lấy luật cũ từ trí nhớ | mọi model |
