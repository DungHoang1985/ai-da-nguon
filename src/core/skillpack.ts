/**
 * Bộ Skill tích hợp sẵn của AI Đa Nguồn.
 * Phỏng theo nguyên tắc của các Skill chính thức của Anthropic — xlsx, docx, pptx (Apache-2.0), dataviz — và plugin legal
 * (review-contract, triage-nda, legal-risk-assessment, compliance-check, signature-request, meeting-briefing, legal-response;
 * Apache-2.0), viết lại bằng tiếng Việt cho thao tác trực tiếp bằng công cụ Office.js của add-in và cho bối cảnh
 * pháp lý – bất động sản Việt Nam. Không chép nguyên văn.
 */
import type { Skill } from "./skills";

/** Hướng dẫn nền luôn đưa vào lời nhắc hệ thống theo ứng dụng (giống chuyên môn sẵn có của add-in chính thức). */
export const CORE_GUIDES: Record<string, string> = {
  Excel: `## Chuẩn làm việc với Excel (luôn áp dụng)
- Dùng công thức, không ghi số tính sẵn: tổng, tỷ lệ, chênh lệch đều là công thức để bảng tự cập nhật.
- Không để lỗi công thức: xem formula_errors sau mỗi lần ghi; trước khi kết thúc với bảng có công thức, chạy excel_find_errors.
- Làm đúng yêu cầu: đúng tên sheet, tên cột, công thức người dùng nêu.
- Sửa tệp có sẵn: theo đúng quy ước của tệp (màu ô nhập liệu, định dạng); chỉ ghi vào ô nhập liệu, không động vào công thức có sẵn.
- Mỗi giả định nằm ở một ô riêng có nhãn, công thức tham chiếu tuyệt đối tới ô đó (=B5*(1+$B$2)), không viết số cứng trong công thức; công thức nhất quán theo hàng/cột; chặn chia cho 0 bằng IFERROR.
- Ghi chú nguồn của số liệu nhập tay (excel_add_comment hoặc ô ghi chú cạnh bảng).
- Mô hình tài chính (khi không có quy ước khác): chữ xanh dương cho số nhập tay, đen cho công thức, xanh lá cho liên kết sang sheet khác, nền vàng cho giả định chính; số "#,##0" có đơn vị ở tiêu đề (vd "Doanh thu (triệu đồng)"), số âm trong ngoặc, số 0 hiện "-", tỷ lệ lưu dạng thập phân định dạng "0.0%", hệ số "0.0x", năm là chữ.
- Không ghi đè dữ liệu có sẵn khi chưa được yêu cầu; dùng vùng trống hoặc sheet mới.
- Biểu đồ: chọn dạng theo mục đích (so sánh → cột/thanh xếp lớn đến nhỏ; xu hướng → đường; tỷ trọng → thanh ngang, biểu đồ tròn chỉ khi được yêu cầu và ≤ 5 phần); một trục tung; màu theo thứ tự màu chủ đề; chú giải khi ≥ 2 chuỗi; tiêu đề nêu ý chính kèm đơn vị (chi tiết: kỹ năng thiet-ke-bieu-do).
- Định dạng có điều kiện: độ lớn → thang một màu nhạt đến đậm; có âm/dương hoặc so với mục tiêu → hai màu với điểm giữa trung tính; màu trạng thái luôn kèm chữ.
- Khi trả lời về số liệu, dẫn địa chỉ ô (vd Sheet1!C12) để người dùng bấm kiểm tra.`,
  Word: `## Chuẩn làm việc với Word (luôn áp dụng)
- Đọc đủ văn bản trước khi kết luận; các điều khoản tác động lẫn nhau.
- Sửa tối thiểu, đúng chỗ: sửa cụm từ bằng word_replace_text, viết lại đoạn bằng word_set_paragraph_text; giữ nguyên kiểu chữ, đánh số, định dạng; nội dung mới theo kiểu đoạn/tiêu đề sẵn có của văn bản.
- Mọi sửa đổi ghi bằng Theo dõi thay đổi để người dùng duyệt; góp ý thì dùng nhận xét neo đúng cụm chữ.
- Không tự bịa số hiệu, điều, khoản văn bản pháp luật; căn cứ nào không có trong tài liệu/tệp đính kèm thì ghi "cần đối chiếu văn bản gốc".
- Kiểm tra nhất quán: thuật ngữ định nghĩa, tham chiếu chéo, đánh số, ngày tháng, số tiền bằng số và bằng chữ, tên các bên.
- Khi nói tới một đoạn, dẫn dạng [¶số đoạn] để người dùng bấm tới.`,
  PowerPoint: `## Chuẩn làm việc với PowerPoint (luôn áp dụng)
- Mỗi slide một thông điệp; bài dài theo cấu trúc kim tự tháp: kết luận chính → tóm tắt → các phần → kết luận/bước tiếp theo.
- Dùng bố cục (layout) và màu, phông của mẫu đang mở; không tự thêm thanh màu trang trí, đường gạch dưới tiêu đề.
- Tiêu đề ngắn (≤ 10 từ), không chấm cuối, cùng vị trí trên các slide nội dung; thân bài 14–16pt trở lên, căn trái; mỗi slide ≤ 6 ý.
- Mỗi slide nên có yếu tố trực quan: biểu đồ gốc (ppt_add_chart) cho số liệu, sơ đồ (ppt_add_process_flow) cho quy trình/mốc thời gian, bảng cho so sánh.
- Không để chữ tràn khỏi khung: nhiều chữ thì tách slide; lề ≥ 36pt, khoảng cách giữa khối đều nhau.
- Kiểm tra lại sau khi tạo: ppt_get_overview để soát lỗi chính tả, thứ tự, chữ mẫu còn sót ("Click to add…", "[cần bổ sung]").
- Biểu đồ: chọn dạng theo mục đích (so sánh → cột/thanh; xu hướng → đường; tỷ trọng → thanh ngang, tròn chỉ khi được yêu cầu và ≤ 5 phần); một trục tung; màu theo màu chủ đề; tiêu đề slide nêu ý chính kèm đơn vị (chi tiết: kỹ năng thiet-ke-bieu-do).
- Khi nói tới slide, ghi "slide số N" để người dùng bấm tới.`,
};

export interface BuiltinSkill extends Skill {
  /** Từ khóa (không dấu, chữ thường) để tự kích hoạt */
  triggers?: string[];
}

export const BUILTIN_SKILLS: BuiltinSkill[] = [
  {
    name: "ra-soat-hop-dong",
    description:
      "Rà soát hợp đồng, phụ lục, thỏa thuận trong Word theo từng điều khoản: xác định loại hợp đồng và bên mình, phân loại sai lệch Xanh/Vàng/Đỏ, ghi nhận xét đúng chỗ, đề xuất câu chữ sửa (Track Changes khi được yêu cầu). Dùng khi người dùng muốn rà soát, góp ý, đánh giá một hợp đồng.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["ra soat hop dong", "gop y hop dong", "review contract", "kiem tra hop dong", "danh gia hop dong", "ra soat phu luc", "ra soat thoa thuan"],
    body: `# Rà soát hợp đồng (phỏng theo plugin legal/review-contract)

1. Xác định loại hợp đồng và **bên mình** (hỏi nếu chưa rõ): bên bán/mua, chủ đầu tư/nhà thầu, bên chuyển nhượng/nhận chuyển nhượng… Kết luận phụ thuộc vào bên mình.
2. Đọc **toàn văn** trước khi kết luận (word_read_paragraphs đến hết; xem cả word_list_comments, word_list_revisions).
3. Phân tích từng nhóm điều khoản: chủ thể và thẩm quyền ký; đối tượng, phạm vi; giá, thanh toán, điều kiện giải ngân; tiến độ, giao nhận, nghiệm thu; chất lượng, bảo hành; quyền và nghĩa vụ; phạt vi phạm, bồi thường thiệt hại, mức trần trách nhiệm; bảo đảm thực hiện; chấm dứt, đơn phương chấm dứt; bất khả kháng; bảo mật; giải quyết tranh chấp, luật áp dụng; hiệu lực, phụ lục, số bản.
4. Phân loại mỗi điểm:
   - **Xanh – chấp nhận**: phù hợp hoặc có lợi hơn cho bên mình.
   - **Vàng – cần đàm phán**: lệch khỏi vị trí mong muốn nhưng thường gặp; nêu phương án dự phòng.
   - **Đỏ – phải xử lý**: rủi ro lớn, thiếu điều khoản cốt lõi, trái quy định; cần ý kiến cấp trên/luật sư.
5. Với mỗi điểm Vàng/Đỏ: word_add_comment neo đúng cụm chữ, theo mẫu "[Vàng/Đỏ] Vấn đề – Rủi ro – Đề xuất câu chữ – Mức ưu tiên (Bắt buộc/Nên có/Có thì tốt)".
6. Chỉ sửa trực tiếp khi được yêu cầu: sửa bằng Track Changes, câu chữ cụ thể sẵn sàng dùng, giọng điệu cân bằng.
7. Căn cứ pháp lý chỉ nêu khi có trong tài liệu hoặc tệp đính kèm; còn lại ghi "cần đối chiếu văn bản gốc".
8. Kết thúc bằng bảng: Điều khoản [¶số đoạn] | Phân loại | Rủi ro | Đề xuất; kèm 3 việc nên làm tiếp.`,
  },
  {
    name: "sang-loc-nda",
    description:
      "Sàng lọc nhanh thỏa thuận bảo mật (NDA) trong Word theo danh mục kiểm tra và xếp loại Xanh (ký được)/Vàng (cần người duyệt)/Đỏ (cần rà soát đầy đủ). Dùng khi người dùng gửi NDA, thỏa thuận bảo mật cần xem nhanh.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["nda", "thoa thuan bao mat", "bao mat thong tin", "cam ket bao mat"],
    body: `# Sàng lọc NDA (phỏng theo plugin legal/triage-nda)

Đọc toàn văn rồi kiểm từng mục, ghi Đạt/Không đạt kèm [¶số đoạn]:
1. Loại NDA: hai chiều hay một chiều, có hợp với quan hệ hợp tác không.
2. Định nghĩa thông tin mật: không quá rộng; có loại trừ chuẩn (đã công khai, đã biết trước, tự phát triển độc lập, nhận hợp pháp từ bên thứ ba, buộc phải cung cấp theo yêu cầu cơ quan nhà nước).
3. Nghĩa vụ bên nhận: mức cẩn trọng hợp lý, chỉ dùng cho mục đích đã nêu, chỉ tiết lộ cho người cần biết và có cam kết tương tự.
4. Thời hạn: thời hạn thỏa thuận và thời hạn bảo mật sau khi chấm dứt hợp lý; tránh nghĩa vụ vô thời hạn với thông tin thông thường.
5. Điều khoản gài thêm: cấm lôi kéo nhân sự, không cạnh tranh, độc quyền, điều khoản "residuals" rộng → đánh dấu.
6. Chế tài: phạt/bồi thường ấn định trước, biện pháp khẩn cấp; luật áp dụng và nơi giải quyết tranh chấp.
7. Hoàn trả/tiêu hủy thông tin có ngoại lệ lưu trữ theo quy định.

Xếp loại: **Xanh** khi mọi mục đạt → ký theo phân quyền; **Vàng** khi có 1–2 điểm lệch thường gặp → chuyển người duyệt kèm điểm cần sửa; **Đỏ** khi có điều khoản gài thêm, thiếu loại trừ chuẩn hoặc chế tài nặng → rà soát đầy đủ. Ghi nhận xét vào từng điểm lệch; cuối cùng báo bảng kết quả và đề xuất hướng xử lý.`,
  },
  {
    name: "danh-gia-rui-ro-phap-ly",
    description:
      "Đánh giá rủi ro pháp lý theo ma trận Mức độ nghiêm trọng × Khả năng xảy ra (1–5), xếp mức Thấp/Trung bình/Cao/Nghiêm trọng, đề xuất biện pháp; lập bảng theo dõi rủi ro trong Excel hoặc bản ghi nhớ trong Word. Dùng khi người dùng muốn đánh giá, chấm điểm, lập danh mục rủi ro của hợp đồng, dự án, thủ tục.",
    hosts: ["Word", "Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["danh gia rui ro", "ma tran rui ro", "risk", "muc do rui ro", "bang rui ro", "danh muc rui ro"],
    body: `# Đánh giá rủi ro pháp lý (phỏng theo plugin legal/legal-risk-assessment)

Chấm mỗi rủi ro hai thang 1–5:
- **Mức độ nghiêm trọng**: 1 Không đáng kể · 2 Thấp · 3 Vừa · 4 Cao · 5 Nghiêm trọng (thiệt hại lớn, đình trệ dự án, chế tài nhà nước).
- **Khả năng xảy ra**: 1 Rất hiếm · 2 Ít khả năng · 3 Có thể · 4 Nhiều khả năng · 5 Gần như chắc chắn.
- **Điểm = Nghiêm trọng × Khả năng**: 1–4 Thấp (Xanh) · 5–9 Trung bình (Vàng) · 10–15 Cao (Cam) · 16–25 Nghiêm trọng (Đỏ).

Với mỗi rủi ro nêu: mô tả, căn cứ (điều khoản/tài liệu nào, [¶] hoặc ô), yếu tố làm tăng/giảm, phương án xử lý (tránh, giảm thiểu, chuyển giao, chấp nhận), rủi ro còn lại, người phụ trách, mốc theo dõi.

- **Trong Excel**: lập bảng theo dõi trên sheet mới "Rui ro": cột STT | Rủi ro | Căn cứ | Nghiêm trọng | Khả năng | Điểm (=D*E bằng công thức) | Mức (=IFS hoặc IF lồng theo điểm) | Biện pháp | Phụ trách | Hạn; tạo Table, định dạng có điều kiện tô màu theo Mức, danh sách chọn 1–5 cho hai cột chấm điểm.
- **Trong Word**: chèn bản ghi nhớ gồm Mô tả – Bối cảnh – Phân tích – Biện pháp – Khuyến nghị – Rủi ro còn lại – Bước tiếp theo, kèm bảng tổng hợp.
- Rủi ro Đỏ hoặc liên quan tranh chấp, cơ quan nhà nước: khuyến nghị xin ý kiến luật sư/cấp có thẩm quyền.`,
  },
  {
    name: "xu-ly-nhan-xet",
    description:
      "Lần lượt xử lý các nhận xét (comment) trong Word: đọc đoạn được neo, sửa đúng chỗ bằng Track Changes, trả lời nhận xét nêu đã sửa gì, đánh dấu đã xử lý khi được yêu cầu. Dùng khi người dùng bảo xử lý, giải quyết, làm theo các nhận xét/góp ý trong văn bản.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["xu ly nhan xet", "giai quyet nhan xet", "lam theo nhan xet", "tra loi nhan xet", "comment", "theo gop y trong van ban"],
    body: `# Xử lý nhận xét trong văn bản

1. word_list_comments: lấy danh sách nhận xét chưa xử lý, đoạn chữ được neo, người viết, các trả lời.
2. Với từng nhận xét (theo thứ tự trong văn bản):
   a. Hiểu yêu cầu; nếu mơ hồ hoặc mâu thuẫn với nhận xét khác thì không sửa, trả lời nhận xét nêu câu hỏi cần làm rõ.
   b. Sửa đúng phần được neo (word_replace_text hoặc word_set_paragraph_text) — Track Changes.
   c. word_reply_comment: tóm tắt ngắn đã sửa gì ("Đã sửa: …").
   d. Chỉ đánh dấu đã xử lý (word_resolve_comment) khi người dùng yêu cầu.
3. Kết thúc: bảng Nhận xét | Đoạn [¶] | Đã làm | Còn vướng.`,
  },
  {
    name: "tom-tat-sua-doi-doi-tac",
    description:
      "Đọc và tóm tắt các sửa đổi được theo dõi (redline) mà đối tác gửi lại trong Word: nhóm theo điều khoản, đánh giá mức độ, đề xuất chấp nhận/từ chối/đàm phán; chấp nhận hoặc từ chối từng sửa đổi khi được yêu cầu. Dùng khi người dùng nhận bản sửa của đối tác, hỏi đối tác đã sửa gì.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["doi tac sua", "ban sua cua doi tac", "redline", "track changes cua", "da sua gi", "sua doi duoc theo doi", "chap nhan sua doi", "tu choi sua doi"],
    body: `# Tóm tắt sửa đổi của đối tác

1. word_list_revisions: lấy toàn bộ sửa đổi (thêm/xóa/định dạng, người sửa, ngày, đoạn [¶]).
2. Nhóm theo điều khoản; với mỗi nhóm nêu: nội dung cũ → mới, ý nghĩa pháp lý/thương mại với bên mình.
3. Phân loại: **Chấp nhận** (câu chữ, không đổi bản chất) · **Đàm phán** (đổi nghĩa vụ, mức tiền, thời hạn) · **Từ chối** (rủi ro lớn, trái lập trường) — nêu lý do ngắn, có thể gửi đối tác.
4. Báo bảng: Điều khoản [¶] | Đối tác sửa | Tác động | Đề xuất.
5. Chỉ chấp nhận/từ chối thật (word_review_revision) khi người dùng ra lệnh; mặc định chỉ báo cáo. Có thể ghi nhận xét phản hồi cho điểm cần đàm phán.`,
  },
  {
    name: "chuan-hoa-bang-tinh",
    description:
      "Làm sạch và chuẩn hóa bảng dữ liệu Excel: dòng tiêu đề, kiểu dữ liệu, ngày tháng, số lưu dạng chữ, khoảng trắng, trùng lặp, định dạng, chuyển thành Table, cố định dòng tiêu đề. Dùng khi người dùng muốn dọn, chuẩn hóa, làm đẹp một bảng.",
    hosts: ["Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["chuan hoa", "lam sach", "don dep bang", "lam dep bang", "dinh dang bang", "clean data"],
    body: `# Chuẩn hóa bảng tính

1. excel_get_overview, excel_read_range toàn vùng; xác định dòng tiêu đề, cột số, cột ngày, cột chữ.
2. Báo vấn đề trước khi sửa: ô trống bất thường, số lưu dạng chữ, ngày sai kiểu, khoảng trắng thừa, dòng trùng (excel_find giúp dò).
3. Mặc định ghi bản sạch sang sheet mới "<tên>_chuan"; chỉ sửa tại chỗ khi được yêu cầu.
4. Định dạng: tiêu đề đậm + tô nền nhạt + viền, số "#,##0", ngày "dd/mm/yyyy", tự giãn cột; tạo Table; cố định dòng tiêu đề (excel_sheet action freeze_rows).
5. Thêm ô kiểm tra (tổng, đếm) bằng công thức để đối chiếu với bảng gốc.
6. Kết thúc: liệt kê thay đổi kèm địa chỉ ô.`,
  },
  {
    name: "kiem-tra-loi-bang-tinh",
    description:
      "Rà soát sổ tính Excel: tìm ô lỗi (#REF!, #DIV/0!, #N/A…), lần nguồn gốc lỗi qua các ô tiền lệ, phát hiện số cứng trong vùng công thức, công thức không nhất quán, tham chiếu sai; đề xuất và sửa khi được đồng ý. Dùng khi người dùng muốn kiểm tra, soát lỗi, audit bảng tính hoặc hỏi một số liệu được tính ra sao.",
    hosts: ["Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["kiem tra loi", "soat loi", "loi cong thuc", "audit", "#ref", "#div", "#n/a", "#value", "tinh ra sao", "lay tu dau", "tinh nhu the nao"],
    body: `# Kiểm tra lỗi bảng tính

1. excel_find_errors cho toàn sổ; với mỗi ô lỗi dùng excel_trace (direction precedents) để lần tới gốc.
2. Đọc các vùng công thức chính (excel_read_range include_formulas) để tìm: số cứng trong hàng công thức, công thức lệch mẫu so với ô bên cạnh, tham chiếu tới ô trống, tham chiếu vòng, đơn vị không khớp.
3. Giải thích một số liệu: lần tiền lệ nhiều cấp, trình bày chuỗi tính từ ô nhập liệu tới kết quả, dẫn địa chỉ từng ô.
4. Báo bảng: Ô | Vấn đề | Nguyên nhân | Cách sửa. Chỉ sửa khi người dùng đồng ý (có Hoàn tác); sửa xong chạy lại excel_find_errors.`,
  },
  {
    name: "mo-hinh-tai-chinh-du-an",
    description:
      "Xây hoặc điền mô hình tài chính dự án (bất động sản) trong Excel: giả định, dòng thời gian, doanh thu, chi phí (đất, xây dựng, chi phí khác), vốn vay, dòng tiền, NPV/IRR, độ nhạy — đúng chuẩn mô hình (màu ô, công thức, kiểm tra). Dùng khi người dùng muốn lập, điền, cập nhật mô hình tài chính, hiệu quả đầu tư, dòng tiền dự án.",
    hosts: ["Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["mo hinh tai chinh", "hieu qua du an", "hieu qua dau tu", "dong tien", "npv", "irr", "phuong an tai chinh", "financial model", "cash flow", "do nhay"],
    body: `# Mô hình tài chính dự án

1. Hỏi/ghi rõ phạm vi: kỳ (tháng/quý/năm), đơn vị tiền, số liệu đầu vào đã có (tệp đính kèm, read_other_app, sổ đang mở). Thiếu số thì để ô nhập liệu nền vàng, ghi "[cần bổ sung]".
2. Cấu trúc sheet: "Gia dinh" (mọi số nhập tay, có nhãn, đơn vị, nguồn) → "Tinh toan" (dòng thời gian, doanh thu, chi phí, vốn vay, thuế, dòng tiền) → "Ket qua" (NPV, IRR, thời gian hoàn vốn, biểu đồ) → "Kiem tra" (cân đối, tổng khớp).
3. Công thức nhất quán theo kỳ; tham chiếu tuyệt đối tới giả định; NPV/IRR dùng =NPV, =IRR (hoặc =XNPV, =XIRR theo ngày).
4. Độ nhạy: bảng hai chiều dựng bằng công thức (không dùng Data Table vì add-in không hỗ trợ) cho 2 biến chính (giá bán, chi phí xây dựng…).
5. Áp dụng quy ước màu và định dạng số của mô hình tài chính; chạy excel_find_errors trước khi báo xong.
6. Báo: các giả định chính, kết quả [ô], các chỗ cần người dùng bổ sung.`,
  },
  {
    name: "phan-tich-du-lieu",
    description:
      "Phân tích dữ liệu trong Excel: tóm tắt, PivotTable, tổng hợp theo nhóm, xu hướng, so sánh, tô màu điểm bất thường bằng định dạng có điều kiện, vẽ biểu đồ và nêu nhận định. Dùng khi người dùng muốn phân tích, tổng hợp, thống kê, báo cáo nhanh từ bảng dữ liệu.",
    hosts: ["Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["phan tich", "tong hop", "thong ke", "pivot", "bao cao nhanh", "xu huong", "so sanh"],
    body: `# Phân tích dữ liệu

1. Đọc cấu trúc dữ liệu (cột, kiểu, số dòng); hỏi câu hỏi kinh doanh nếu chưa rõ.
2. Tổng hợp bằng excel_create_pivot (dữ liệu lớn) hoặc công thức SUMIFS/COUNTIFS trên sheet mới "Phan tich".
3. Làm nổi bật: excel_conditional_format (thang màu, thanh dữ liệu, ngưỡng) cho điểm bất thường.
4. Vẽ biểu đồ phù hợp: xu hướng → đường; so sánh nhóm → cột; tỷ trọng (≤ 6 nhóm) → tròn.
5. Nêu 3–5 nhận định có số liệu, dẫn địa chỉ ô; chỉ ra hạn chế của dữ liệu.`,
  },
  {
    name: "bao-cao-tien-do-du-an",
    description:
      "Dựng bài trình chiếu báo cáo tiến độ dự án (bất động sản) trong PowerPoint: bìa, tổng quan, tiến độ pháp lý, tiến độ thi công, tài chính, vướng mắc – kiến nghị, kế hoạch tiếp theo. Dùng khi người dùng muốn tạo hoặc cập nhật slide báo cáo tiến độ.",
    hosts: ["PowerPoint"],
    source: "builtin",
    enabled: true,
    triggers: ["bao cao tien do", "tien do du an", "bao cao du an", "progress report"],
    body: `# Báo cáo tiến độ dự án

1. ppt_get_overview: xem bố cục và slide đã có; dùng đúng mẫu của tệp.
2. Nguồn số liệu: tệp đính kèm, nội dung chat, read_other_app (Excel/Word đang mở khung AI Đa Nguồn). Không bịa số; thiếu thì ghi "[cần bổ sung]".
3. Cấu trúc mặc định: Bìa → Tổng quan dự án → Tiến độ pháp lý (bảng: Thủ tục | Cơ quan | Trạng thái | Mốc) → Tiến độ thi công (ppt_add_chart % hoàn thành) → Tài chính → Vướng mắc & kiến nghị → Kế hoạch kỳ tới (ppt_add_process_flow).
4. Mỗi slide một thông điệp, ≤ 6 ý; số liệu dùng biểu đồ gốc; quy trình dùng sơ đồ.
5. Kết thúc: liệt kê slide đã tạo và các chỗ "[cần bổ sung]".`,
  },
  {
    name: "trinh-chieu-tu-tai-lieu",
    description:
      "Chuyển tài liệu (tệp đính kèm PDF/DOCX, văn bản Word hoặc số liệu Excel đang mở) thành bài trình chiếu PowerPoint có cấu trúc: thông điệp chính, tóm tắt, các phần, kết luận và bước tiếp theo. Dùng khi người dùng muốn làm slide từ một tài liệu, báo cáo, văn bản.",
    hosts: ["PowerPoint"],
    source: "builtin",
    enabled: true,
    triggers: ["tu tai lieu", "tu van ban", "tu file", "tu tep", "lam slide tu", "tao slide tu", "chuyen thanh slide", "tom tat thanh slide"],
    body: `# Trình chiếu từ tài liệu

1. Lấy nội dung: tệp đính kèm trong câu hỏi, hoặc read_other_app("Word"/"Excel").
2. Xác định mục đích (ra quyết định, báo cáo, thuyết trình) và người nghe; chọn số slide (≈ 1,5–2 phút/slide nếu trình bày trực tiếp).
3. Dàn ý kim tự tháp: thông điệp chính → tóm tắt điều hành → mỗi phần một nhóm slide → kết luận và bước tiếp theo. Báo dàn ý ngắn rồi tạo luôn (không chờ trừ khi người dùng muốn duyệt).
4. Tạo slide bằng ppt_add_slide theo bố cục mẫu; số liệu → ppt_add_chart; quy trình/mốc → ppt_add_process_flow; so sánh → ppt_add_table.
5. Dẫn nguồn (trang/điều/ô) ở chú thích nhỏ khi dùng số liệu; soát lại bằng ppt_get_overview.`,
  },
  {
    name: "kiem-tra-bai-trinh-chieu",
    description:
      "Soát bài trình chiếu PowerPoint: chữ mẫu còn sót, lỗi chính tả, tiêu đề lệch vị trí/kiểu, slide quá nhiều chữ, cỡ chữ quá nhỏ, thứ tự và mạch nội dung, số liệu không khớp; sửa khi được đồng ý. Dùng khi người dùng muốn kiểm tra, soát, hoàn thiện bài trình chiếu trước khi gửi.",
    hosts: ["PowerPoint"],
    source: "builtin",
    enabled: true,
    triggers: ["kiem tra bai trinh chieu", "soat slide", "ra soat slide", "hoan thien slide", "deck check", "kiem tra slide"],
    body: `# Kiểm tra bài trình chiếu

1. ppt_get_overview toàn bài.
2. Soát: chữ mẫu còn sót ("Click to add…", "Lorem", "[cần bổ sung]"); lỗi chính tả; tiêu đề có dấu chấm cuối hoặc khác vị trí; slide > 6 ý hoặc chữ dài; số liệu mâu thuẫn giữa các slide; thứ tự kể chuyện; slide thiếu yếu tố trực quan.
3. Báo bảng: Slide số | Vấn đề | Đề xuất.
4. Sửa khi được đồng ý: ppt_set_shape_text, ppt_update_shape, ppt_move_slide; có Hoàn tác.`,
  },
  {
    name: "kiem-tra-tuan-thu",
    description:
      "Kiểm tra tuân thủ một việc định làm (phương án, chương trình bán hàng, hồ sơ, văn bản, quy trình) trước khi triển khai: quy định nào áp dụng, yêu cầu đã đáp ứng chưa, rủi ro, phê duyệt cần có, kết luận Làm được / Làm được có điều kiện / Cần xem xét thêm. Dùng khi người dùng hỏi 'có đúng quy định không', 'cần xin phép gì', 'điều kiện pháp lý để…', 'kiểm tra tuân thủ'.",
    hosts: ["Word", "Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["kiem tra tuan thu", "tuan thu", "compliance", "dung quy dinh", "can xin phep", "can phe duyet", "dieu kien phap ly", "co hop phap", "co vi pham"],
    body: `# Kiểm tra tuân thủ (phỏng theo plugin legal/compliance-check)

1. Làm rõ việc định làm: ai làm, làm gì, ở đâu, khi nào, ai bị tác động (người mua nhà, khách hàng, cơ quan nhà nước, dữ liệu cá nhân…). Thiếu thông tin quan trọng thì hỏi một lần, gom các câu hỏi.
2. Đọc tài liệu đang mở và tệp đính kèm; số liệu ở ứng dụng khác thì read_other_app.
3. Xác định nhóm quy định có thể áp dụng: đầu tư, đất đai, quy hoạch, xây dựng, nhà ở, kinh doanh bất động sản, môi trường, phòng cháy chữa cháy, thuế, bảo vệ dữ liệu cá nhân, quảng cáo, cạnh tranh, quy chế nội bộ. CHỈ nêu số hiệu, điều, khoản khi có trong tài liệu hoặc tệp đính kèm; còn lại ghi "cần đối chiếu văn bản gốc" — không trích từ trí nhớ.
4. Trình bày kết quả:
   - Kết luận nhanh: Làm được / Làm được có điều kiện / Cần xem xét thêm.
   - Bảng Quy định áp dụng | Vì sao áp dụng | Yêu cầu chính | Nguồn ([¶n], tên tệp, hoặc "cần đối chiếu").
   - Bảng Yêu cầu | Tình trạng (Đã đáp ứng / Chưa / Chưa rõ) | Việc cần làm.
   - Bảng Rủi ro | Mức (Cao / Trung bình / Thấp) | Cách giảm thiểu.
   - Tối đa 5 việc nên làm, theo thứ tự ưu tiên.
   - Phê duyệt hoặc ý kiến cần có: Ai | Vì sao | Trạng thái.
   - Phần nên có chuyên gia hoặc luật sư xem thêm.
5. Word: trả lời trong khung chat; người dùng muốn đưa vào văn bản thì word_insert_paragraphs và word_insert_table ở cuối văn bản (ghi bằng Theo dõi thay đổi). Excel: tạo sheet "KiemTraTuanThu" chứa các bảng trên (excel_write_range, excel_create_table), cột Tình trạng có danh sách chọn (excel_data_validation).`,
  },
  {
    name: "kiem-tra-truoc-ky",
    description:
      "Soát hợp đồng hoặc văn bản trong Word trước khi trình ký: còn sửa đổi chưa duyệt, nhận xét chưa xử lý, phụ lục được nhắc tới có đủ không, tên pháp nhân, người ký, chức danh, ủy quyền, ngày tháng, số tiền bằng số và bằng chữ, đánh số và tham chiếu, chữ mẫu còn sót. Kết luận ĐẠT hoặc CÓ VẤN ĐỀ kèm danh sách cần sửa. Dùng khi người dùng nói 'trước khi ký', 'trình ký', 'chốt bản ký', 'kiểm tra bản cuối'.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["truoc khi ky", "trinh ky", "ban ky", "chot ban", "ban cuoi", "san sang ky", "ky ket", "signature"],
    body: `# Kiểm tra trước khi ký (phỏng theo plugin legal/signature-request, phần kiểm tra trước ký)

1. word_list_revisions: còn sửa đổi chưa chấp nhận hoặc từ chối → CÓ VẤN ĐỀ, vì bản ký phải là bản sạch. Không tự chấp nhận tất cả; hỏi người dùng.
2. word_list_comments: còn nhận xét chưa xử lý → liệt kê.
3. Đọc toàn văn (word_read_paragraphs) và kiểm:
   - Tên pháp nhân, mã số doanh nghiệp, địa chỉ, người đại diện, chức danh giống nhau ở phần chủ thể, phần ký và mọi chỗ nhắc lại.
   - Người ký có thẩm quyền: là người đại diện theo pháp luật, hoặc có văn bản ủy quyền được dẫn chiếu (số, ngày).
   - Phụ lục, biểu mẫu, bản vẽ được nhắc tới đều có trong văn bản hoặc ghi rõ là kèm theo; số thứ tự phụ lục liên tục.
   - Ngày tháng: ngày ký để trống hoặc đúng; thời hạn và mốc tiến độ không mâu thuẫn.
   - Số tiền: bằng số khớp bằng chữ; đơn vị tiền tệ; đã gồm hay chưa gồm thuế ghi nhất quán.
   - Đánh số điều, khoản; tham chiếu chéo trỏ đúng; thuật ngữ định nghĩa dùng nhất quán.
   - Số bản, ngôn ngữ ưu tiên, hiệu lực; ký nháy từng trang nếu áp dụng.
   - Chữ mẫu còn sót: "…", "[ ]", "xxx", "[cần bổ sung]", chỗ tô vàng.
4. Kết quả: "Kiểm tra trước ký: ĐẠT" hoặc "CÓ VẤN ĐỀ", kèm bảng Hạng mục | Kết quả | Vị trí [¶n] | Cần làm. Hỏi người dùng có muốn đặt nhận xét tại từng chỗ (word_add_comment) không.
5. Add-in không gửi văn bản đi ký điện tử; nhắc lưu bản đã ký vào hồ sơ.`,
  },
  {
    name: "chuan-bi-hop",
    description:
      "Chuẩn bị tài liệu cho cuộc họp có nội dung pháp lý (đàm phán hợp đồng, làm việc với cơ quan nhà nước, họp ban dự án, họp đối tác): tổng hợp từ văn bản đang mở, tệp đính kèm và ứng dụng khác thành bản chuẩn bị gồm bối cảnh, vấn đề còn mở, lưu ý pháp lý, ý chính cần nói, câu hỏi cần đặt, quyết định cần có, giới hạn không nhượng bộ, thông tin còn thiếu. Dùng khi người dùng nói 'chuẩn bị họp', 'tài liệu họp', 'briefing', 'mai đi đàm phán'.",
    hosts: ["Word", "PowerPoint"],
    source: "builtin",
    enabled: true,
    triggers: ["chuan bi hop", "tai lieu hop", "briefing", "cuoc hop", "buoi hop", "di hop", "lam viec voi so", "hop voi", "chuan bi dam phan"],
    body: `# Chuẩn bị họp (phỏng theo plugin legal/meeting-briefing)

1. Xác định: tên và loại cuộc họp, thời gian, thành phần và vai trò của họ, vai trò của mình (trình bày, tư vấn, đàm phán, dự họp), chương trình. Thiếu thì hỏi gọn một lần.
2. Thu thập: văn bản đang mở, tệp đính kèm, số liệu Excel hoặc slide PowerPoint qua read_other_app. Add-in không đọc lịch và email, nên thông tin đó lấy từ người dùng.
3. Viết bản chuẩn bị theo các mục:
   - Thông tin cuộc họp · Thành phần (tên, vị trí, mối quan tâm) · Chương trình.
   - Bối cảnh (3–5 câu) · Tài liệu chính (dẫn [¶n], Sheet!A1, slide số N).
   - Vấn đề còn mở · Lưu ý pháp lý (căn cứ chỉ nêu khi có trong tài liệu; còn lại "cần đối chiếu văn bản gốc").
   - Ý chính cần nói (3–5) · Câu hỏi cần đặt · Quyết định cần có · Giới hạn không nhượng bộ.
   - Thông tin còn thiếu hoặc chưa kiểm chứng.
   - Họp đàm phán: thêm lập trường dự kiến của đối tác và phương án dự phòng. Làm việc với cơ quan nhà nước: hồ sơ đã nộp, thời hạn, văn bản trả lời trước đó.
4. Word: chèn bản chuẩn bị ở cuối văn bản (word_insert_paragraphs, đề mục dạng tiêu đề). PowerPoint: 3–5 slide (bối cảnh, vấn đề, phương án, quyết định cần có).
5. Sau họp: người dùng dán biên bản hoặc ghi chú thì lập bảng Việc | Người phụ trách | Hạn | Trạng thái.`,
  },
  {
    name: "soan-phan-hoi-phap-ly",
    description:
      "Soạn văn bản trả lời một yêu cầu hoặc câu hỏi pháp lý thường gặp (khách hàng hỏi về pháp lý, tiến độ dự án; đối tác đề nghị ký NDA; yêu cầu cung cấp thông tin hoặc dữ liệu cá nhân; văn bản đôn đốc; khiếu nại), có bước kiểm tra các dấu hiệu không được trả lời theo mẫu mà phải chuyển cấp trên hoặc luật sư. Dùng khi người dùng nói 'soạn văn bản trả lời', 'trả lời công văn', 'phản hồi khách hàng', 'trả lời khiếu nại'.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["van ban tra loi", "tra loi cong van", "cong van tra loi", "phan hoi khach hang", "tra loi yeu cau", "phan hoi yeu cau", "tra loi khieu nai", "thu tra loi"],
    body: `# Soạn phản hồi pháp lý (phỏng theo plugin legal/legal-response)

1. Xác định loại yêu cầu và người gửi; đọc văn bản đến (đang mở hoặc đính kèm).
2. Kiểm tra dấu hiệu phải chuyển cấp trên hoặc luật sư: có kiện tụng hoặc đe dọa kiện; yêu cầu của tòa án, cơ quan điều tra, thanh tra, cơ quan quản lý; khiếu nại đông người hoặc có yếu tố báo chí; số tiền lớn hoặc nội dung có thể bị coi là thừa nhận trách nhiệm; dữ liệu cá nhân nhạy cảm; thời hạn trả lời sắp hết. Có dấu hiệu thì nói rõ lý do, đề xuất người duyệt; nếu người dùng vẫn cần bản nháp thì ghi đầu văn bản "BẢN NHÁP – CẦN DUYỆT".
3. Hỏi gọn chi tiết còn thiếu: số và ngày văn bản đến, nội dung được phép cam kết, người ký, hạn trả lời.
4. Soạn: dẫn chiếu văn bản đến; trả lời từng ý theo thứ tự câu hỏi; không thừa nhận điều chưa xác nhận; không hứa ngoài thẩm quyền; căn cứ pháp lý chỉ nêu khi có trong tài liệu, còn lại ghi "cần đối chiếu văn bản gốc"; giọng lịch sự, rõ ràng; kết bằng đầu mối liên hệ.
5. Văn bản gửi cơ quan nhà nước: theo thể thức văn bản hành chính hiện hành (quốc hiệu, số và ký hiệu, trích yếu, nơi nhận), dùng kiểu đoạn sẵn có của văn bản.
6. Chèn vào Word bằng word_insert_paragraphs (Theo dõi thay đổi); báo các chỗ cần điền [ ].`,
  },
  {
    name: "thiet-ke-bieu-do",
    description:
      "Chọn dạng biểu đồ và trình bày biểu đồ cho đúng trong Excel và PowerPoint: chọn theo việc người xem cần (so sánh, xu hướng, cơ cấu, chênh lệch), màu theo thứ tự màu chủ đề, một trục tung, chú giải và nhãn số chọn lọc, tiêu đề nêu ý chính. Dùng khi người dùng muốn vẽ, sửa, làm đẹp biểu đồ hoặc hỏi nên dùng biểu đồ gì.",
    hosts: ["Excel", "PowerPoint"],
    source: "builtin",
    enabled: true,
    triggers: ["bieu do", "do thi", "chart", "ve bieu", "graph", "truc quan"],
    body: `# Thiết kế biểu đồ (phỏng theo nguyên tắc của skill dataviz)

1. Chọn dạng trước, chọn màu sau. Chọn theo việc người xem cần làm:
   - So sánh độ lớn giữa các mục → cột; thanh ngang nếu tên dài hoặc trên 8 mục; xếp từ lớn đến nhỏ.
   - Xu hướng theo thời gian → đường (vùng nếu chỉ 1 chuỗi).
   - Cơ cấu, tỷ trọng → thanh ngang xếp hạng; cột chồng khi so cơ cấu qua nhiều kỳ. Biểu đồ tròn hoặc vành khuyên chỉ khi người dùng yêu cầu và không quá 5 phần.
   - Chênh lệch so với kế hoạch hoặc kỳ trước → cột âm dương quanh mốc 0.
   - Chỉ một con số chính → ô số lớn kèm mức tăng giảm, không vẽ biểu đồ một cột.
2. Màu: theo thứ tự màu chủ đề (accent 1 → 6), mỗi đối tượng giữ một màu ở mọi biểu đồ. Trên 6 chuỗi → gộp nhóm nhỏ thành "Khác" hoặc tách nhiều biểu đồ. Muốn nhấn một chuỗi thì chuỗi đó dùng màu chủ đề, còn lại màu xám. Không dùng dải cầu vồng. Thể hiện độ lớn → một màu từ nhạt đến đậm; có âm và dương → hai màu với điểm giữa trung tính.
3. Một trục tung. Hai đại lượng khác đơn vị (tiền và %) → hai biểu đồ riêng.
4. Có chú giải khi từ 2 chuỗi trở lên (đặt phía dưới); 1 chuỗi thì không cần, tiêu đề nêu tên. Nhãn số chọn lọc (giá trị cuối, lớn nhất, mốc quan trọng), không ghi số trên mọi điểm. Đường lưới mờ hoặc bỏ.
5. Tiêu đề nêu ý chính và đơn vị, ví dụ "Doanh thu 2026 tăng 32% (tỷ đồng)". Trục của cột và thanh bắt đầu từ 0.
6. Nhãn và chú giải dùng màu chữ thường, không dùng màu của chuỗi. Màu trạng thái (đạt, cảnh báo, không đạt) chỉ dùng cho trạng thái và luôn kèm chữ.
7. Excel: excel_create_chart từ vùng dữ liệu có dòng tiêu đề. PowerPoint: ppt_add_chart (biểu đồ gốc, sửa được dữ liệu). Xong thì xem lại bằng công cụ tổng quan.`,
  },
];

function norm(s: string): string {
  return String(s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "d")
    .toLowerCase();
}

/** Chọn Skill tự kích hoạt theo nội dung câu hỏi (tối đa 2), dựa trên từ khóa. */
export function autoSelectSkills(message: string, skills: (Skill & { triggers?: string[] })[], max = 2): Skill[] {
  const text = norm(message);
  const scored = skills
    .map((s) => {
      const triggers = s.triggers?.length ? s.triggers : [norm(s.name.replace(/-/g, " "))];
      const hits = triggers.filter((t) => text.includes(norm(t)));
      return { s, score: hits.reduce((n, t) => n + t.split(" ").length, 0) };
    })
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score);
  return scored.slice(0, max).map((x) => x.s);
}
