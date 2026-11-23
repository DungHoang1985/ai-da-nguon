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
- Dùng công thức (tên hàm tiếng Anh, dấu phẩy), không ghi số tính sẵn: tổng, tỷ lệ, chênh lệch đều là công thức để bảng tự cập nhật.
- Không để lỗi công thức: trước khi kết thúc với bảng có công thức, chạy excel_find_errors.
- Làm đúng yêu cầu: đúng tên sheet, tên cột, công thức người dùng nêu.
- Không ghi đè, không xóa dữ liệu có sẵn khi chưa được yêu cầu; cần chỗ mới thì dùng vùng trống hoặc sheet mới. Sửa tệp có sẵn: theo đúng quy ước của tệp (màu ô nhập liệu, định dạng); chỉ ghi vào ô nhập liệu, không động vào công thức có sẵn.
- Mỗi giả định nằm ở một ô riêng có nhãn, công thức tham chiếu tuyệt đối tới ô đó (=B5*(1+$B$2)), không viết số cứng trong công thức; công thức nhất quán theo hàng/cột; chặn chia cho 0 bằng IFERROR.
- Ghi chú nguồn của số liệu nhập tay (excel_add_comment hoặc ô ghi chú cạnh bảng).
- Mô hình tài chính (khi không có quy ước khác): chữ xanh dương cho số nhập tay, đen cho công thức, xanh lá cho liên kết sang sheet khác, nền vàng cho giả định chính; số "#,##0" có đơn vị ở tiêu đề (vd "Doanh thu (triệu đồng)"), số âm trong ngoặc, số 0 hiện "-", tỷ lệ lưu dạng thập phân định dạng "0.0%", hệ số "0.0x", năm là chữ.
- Biểu đồ: chọn dạng theo mục đích (so sánh → cột/thanh xếp lớn đến nhỏ; xu hướng → đường; tỷ trọng → thanh ngang, biểu đồ tròn chỉ khi được yêu cầu và ≤ 5 phần); một trục tung; màu theo thứ tự màu chủ đề; chú giải khi ≥ 2 chuỗi; tiêu đề nêu ý chính kèm đơn vị (chi tiết: kỹ năng thiet-ke-bieu-do).
- Định dạng có điều kiện: độ lớn → thang một màu nhạt đến đậm; có âm/dương hoặc so với mục tiêu → hai màu với điểm giữa trung tính; màu trạng thái luôn kèm chữ.`,
  Word: `## Chuẩn làm việc với Word (luôn áp dụng)
- Rà soát toàn diện thì đọc đủ văn bản trước khi kết luận, vì các điều khoản tác động lẫn nhau.
- Sửa tối thiểu, đúng chỗ, giữ nguyên kiểu chữ, đánh số, định dạng; góp ý thì dùng nhận xét neo đúng cụm chữ.
- Căn cứ pháp lý không có trong tài liệu hoặc tệp đính kèm thì ghi "cần đối chiếu văn bản gốc".
- Kiểm tra nhất quán: thuật ngữ định nghĩa, tham chiếu chéo, đánh số, ngày tháng, số tiền bằng số và bằng chữ, tên các bên.
- Soạn văn bản hành chính theo kỹ năng soan-thao-van-ban-hanh-chinh.`,
  PowerPoint: `## Chuẩn làm việc với PowerPoint (luôn áp dụng)
- Mỗi slide một thông điệp; bài dài theo cấu trúc kim tự tháp: kết luận chính → tóm tắt → các phần → kết luận/bước tiếp theo.
- Dùng bố cục (layout) và màu, phông của mẫu đang mở; không tự thêm thanh màu trang trí, đường gạch dưới tiêu đề.
- Tiêu đề ngắn (≤ 10 từ), không chấm cuối, cùng vị trí trên các slide nội dung; thân bài 14–16pt trở lên, căn trái; mỗi slide ≤ 6 ý.
- Mỗi slide nên có yếu tố trực quan: biểu đồ cho số liệu, sơ đồ cho quy trình/mốc thời gian, bảng cho so sánh.
- Không để chữ tràn khỏi khung: nhiều chữ thì tách slide; hình và ô chữ không chồng nhau, lề ≥ 36pt, khoảng cách giữa khối đều nhau.
- Soát lại khi xong: lỗi chính tả, thứ tự, chữ mẫu còn sót ("Click to add…", "[cần bổ sung]").
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
    brief: "Rà soát hợp đồng theo điều khoản, phân loại Xanh/Vàng/Đỏ, ghi nhận xét, đề xuất câu chữ. Dùng khi rà soát, góp ý, đánh giá hợp đồng.",
    description:
      "Rà soát hợp đồng, phụ lục, thỏa thuận trong Word theo từng điều khoản: xác định loại hợp đồng và bên mình, phân loại sai lệch Xanh/Vàng/Đỏ, ghi nhận xét đúng chỗ, đề xuất câu chữ sửa (Track Changes khi được yêu cầu). Dùng khi người dùng muốn rà soát, góp ý, đánh giá một hợp đồng.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["ra soat hop dong", "ra soat mau hop dong", "kiem tra mau hop dong", "gop y mau hop dong", "gop y hop dong", "review contract", "kiem tra hop dong", "danh gia hop dong", "ra soat phu luc", "ra soat thoa thuan"],
    body: `# Rà soát hợp đồng (phỏng theo plugin legal/review-contract)

1. Xác định loại hợp đồng và **bên mình** (hỏi nếu chưa rõ): bên bán/mua, chủ đầu tư/nhà thầu, bên chuyển nhượng/nhận chuyển nhượng… Kết luận phụ thuộc vào bên mình.
2. Xem mục lục điều khoản (word_get_outline) rồi đọc **toàn văn** trước khi kết luận (word_read_paragraphs, gọi tiếp theo next_start đến hết; xem cả word_list_comments, word_list_revisions).
3. Phân tích từng nhóm điều khoản: chủ thể và thẩm quyền ký; đối tượng, phạm vi; giá, thanh toán, điều kiện giải ngân; tiến độ, giao nhận, nghiệm thu; chất lượng, bảo hành; quyền và nghĩa vụ; phạt vi phạm, bồi thường thiệt hại, mức trần trách nhiệm; bảo đảm thực hiện; chấm dứt, đơn phương chấm dứt; bất khả kháng; bảo mật; giải quyết tranh chấp, luật áp dụng; hiệu lực, phụ lục, số bản.
4. Phân loại mỗi điểm:
   - **Xanh – chấp nhận**: phù hợp hoặc có lợi hơn cho bên mình.
   - **Vàng – cần đàm phán**: lệch khỏi vị trí mong muốn nhưng thường gặp; nêu phương án dự phòng.
   - **Đỏ – phải xử lý**: rủi ro lớn, thiếu điều khoản cốt lõi, trái quy định; cần ý kiến cấp trên/luật sư.
5. Với mỗi điểm Vàng/Đỏ: word_add_comment neo đúng cụm chữ, theo mẫu "[Vàng/Đỏ] Vấn đề – Rủi ro – Đề xuất câu chữ – Mức ưu tiên (Bắt buộc/Nên có/Có thì tốt)".
6. Chỉ sửa trực tiếp khi được yêu cầu: sửa bằng Track Changes, câu chữ cụ thể sẵn sàng dùng, giọng điệu cân bằng.
7. Căn cứ pháp lý chỉ nêu khi có trong tài liệu hoặc tệp đính kèm; còn lại ghi "cần đối chiếu văn bản gốc".
8. Kết thúc bằng bảng: Điều khoản (số điều, [¶số đoạn]) | Phân loại | Rủi ro | Đề xuất; kèm 3 việc nên làm tiếp.`,
  },
  {
    name: "sang-loc-nda",
    brief: "Sàng lọc nhanh NDA, xếp loại Xanh/Vàng/Đỏ. Dùng khi có thỏa thuận bảo mật cần xem nhanh.",
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
    brief: "Chấm rủi ro pháp lý theo ma trận mức độ × khả năng, đề xuất biện pháp. Dùng khi đánh giá, lập danh mục rủi ro.",
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
    brief: "Xử lý từng nhận xét (comment): sửa đúng chỗ, trả lời, đánh dấu xong. Dùng khi được nhờ xử lý, làm theo nhận xét, góp ý.",
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
   b. Sửa đúng phần được neo: word_replace_text kèm paragraph_index của đoạn được neo (đổi một cụm), hoặc word_set_paragraph_text kèm expect (viết lại câu) — Track Changes. Văn bản có thể đang có sửa đổi của người khác: chỉ sửa đúng phần nhận xét yêu cầu.
   c. word_reply_comment: tóm tắt ngắn đã sửa gì ("Đã sửa: …").
   d. Chỉ đánh dấu đã xử lý (word_resolve_comment) khi người dùng yêu cầu.
3. Kết thúc: bảng Nhận xét | Điều, đoạn [¶] | Đã làm | Còn vướng.`,
  },
  {
    name: "tom-tat-sua-doi-doi-tac",
    brief: "Tóm tắt redline của đối tác theo điều khoản, đề xuất chấp nhận/từ chối. Dùng khi hỏi đối tác đã sửa gì.",
    description:
      "Đọc và tóm tắt các sửa đổi được theo dõi (redline) mà đối tác gửi lại trong Word: nhóm theo điều khoản, đánh giá mức độ, đề xuất chấp nhận/từ chối/đàm phán; chấp nhận hoặc từ chối từng sửa đổi khi được yêu cầu. Dùng khi người dùng nhận bản sửa của đối tác, hỏi đối tác đã sửa gì.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["doi tac sua", "ban sua cua doi tac", "redline", "track changes cua", "da sua gi", "sua doi duoc theo doi", "chap nhan sua doi", "tu choi sua doi"],
    body: `# Tóm tắt sửa đổi của đối tác

1. word_list_revisions: lấy TOÀN BỘ sửa đổi — kết quả có next_start thì gọi tiếp cho tới khi số đã đọc bằng total. Mỗi nhóm là một đoạn, có số điều (clause), câu trước khi sửa (before), câu sau khi sửa (after) và từng sửa đổi (thêm/xóa/định dạng, người sửa, ngày).
2. Nhóm theo điều khoản; với mỗi nhóm nêu: nội dung cũ → mới, ý nghĩa pháp lý/thương mại với bên mình.
3. Phân loại: **Chấp nhận** (câu chữ, không đổi bản chất) · **Đàm phán** (đổi nghĩa vụ, mức tiền, thời hạn) · **Từ chối** (rủi ro lớn, trái lập trường) — nêu lý do ngắn, có thể gửi đối tác.
4. Báo bảng: Điều khoản (số điều, [¶]) | Đối tác sửa | Tác động | Đề xuất; nêu tổng số sửa đổi đã đọc.
5. Chỉ chấp nhận/từ chối thật (word_review_revision) khi người dùng ra lệnh; mặc định chỉ báo cáo. Có thể ghi nhận xét phản hồi cho điểm cần đàm phán.`,
  },
  {
    name: "chuan-hoa-bang-tinh",
    brief: "Dọn, chuẩn hóa bảng dữ liệu: tiêu đề, kiểu dữ liệu, ngày, số dạng chữ, trùng lặp, định dạng. Dùng khi muốn dọn, chuẩn hóa, làm đẹp bảng.",
    description:
      "Làm sạch và chuẩn hóa bảng dữ liệu Excel: dòng tiêu đề, kiểu dữ liệu, ngày tháng, số lưu dạng chữ, khoảng trắng, trùng lặp, định dạng, chuyển thành Table, cố định dòng tiêu đề. Dùng khi người dùng muốn dọn, chuẩn hóa, làm đẹp một bảng.",
    hosts: ["Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["chuan hoa", "lam sach", "don dep bang", "lam dep bang", "dinh dang bang", "clean data"],
    body: `# Chuẩn hóa bảng tính

1. excel_get_overview, excel_read_range toàn vùng; xác định dòng tiêu đề, cột số, cột ngày, cột chữ.
2. Báo vấn đề trước khi sửa: ô trống bất thường, số lưu dạng chữ, ngày sai kiểu, khoảng trắng thừa, dòng trùng (excel_find giúp dò).
3. Người dùng nêu tên sheet cần làm sạch ("làm sạch sheet X") thì sửa ngay trên sheet đó (có Hoàn tác) và xóa phần thừa của bảng cũ; chỉ ghi bản sạch sang sheet mới "<tên>_chuan" khi người dùng không nói rõ chỗ hoặc muốn giữ bản gốc.
4. Định dạng: tiêu đề đậm + tô nền nhạt + viền, số "#,##0", ngày "dd/mm/yyyy", tự giãn cột; tạo Table; cố định dòng tiêu đề (excel_sheet action freeze_rows).
5. Thêm ô kiểm tra (tổng, đếm) bằng công thức để đối chiếu với bảng gốc.
6. Kết thúc: liệt kê thay đổi kèm địa chỉ ô.`,
  },
  {
    name: "kiem-tra-loi-bang-tinh",
    brief: "Soát lỗi sổ tính: ô lỗi, nguồn gốc lỗi, số gõ cứng, công thức lệch. Dùng khi kiểm tra, audit bảng tính hoặc hỏi số được tính ra sao.",
    description:
      "Rà soát sổ tính Excel: tìm ô lỗi (#REF!, #DIV/0!, #N/A…), lần nguồn gốc lỗi qua các ô tiền lệ, phát hiện số cứng trong vùng công thức, công thức không nhất quán, tham chiếu sai; đề xuất và sửa khi được đồng ý. Dùng khi người dùng muốn kiểm tra, soát lỗi, audit bảng tính hoặc hỏi một số liệu được tính ra sao.",
    hosts: ["Excel"],
    source: "builtin",
    enabled: true,
    triggers: ["kiem tra loi", "soat loi", "loi cong thuc", "audit", "#ref", "#div", "#n/a", "#value", "tinh ra sao", "lay tu dau", "tinh nhu the nao"],
    body: `# Kiểm tra lỗi bảng tính

1. excel_find_errors cho toàn sổ; với mỗi ô lỗi dùng excel_trace (direction precedents) để lần tới gốc.
2. Soát số gõ cứng, công thức lệch mẫu so với ô bên cạnh, tỷ lệ gõ trong công thức (*0.03), liên kết tệp ngoài: gọi excel_audit_model trước (soát bằng mã, không bỏ sót ô), rồi đọc các ô nó nêu (excel_read_range include_formulas) để xác nhận và tìm thêm tham chiếu lệch cột, vùng SUM thiếu ô, tham chiếu vòng, đơn vị không khớp.
3. Chỉ báo ô có sai rõ ràng, nêu được cách sửa cụ thể. Không báo ô trống, tiêu đề, ô "nên kiểm tra lại", và không nhắc lại một ô hai lần.
4. Giải thích một số liệu: lần tiền lệ nhiều cấp, trình bày chuỗi tính từ ô nhập liệu tới kết quả, dẫn địa chỉ từng ô.
5. Báo bảng: Ô | Vấn đề | Nguyên nhân | Cách sửa. Chỉ sửa khi người dùng đồng ý (có Hoàn tác); sửa xong chạy lại excel_find_errors.`,
  },
  {
    name: "mo-hinh-tai-chinh-du-an",
    brief: "Lập, điền mô hình tài chính dự án BĐS: giả định, doanh thu, chi phí, vốn vay, dòng tiền, NPV/IRR, độ nhạy. Dùng khi nói mô hình tài chính, hiệu quả đầu tư.",
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
5. Áp dụng quy ước màu và định dạng số của mô hình tài chính — định dạng cả mô hình trong MỘT lệnh excel_format_range (vùng khác kiểu đưa vào more), không định dạng từng hàng; chạy excel_find_errors trước khi báo xong.
6. Báo: các giả định chính, kết quả [ô], các chỗ cần người dùng bổ sung.`,
  },
  {
    name: "phan-tich-du-lieu",
    brief: "Phân tích bảng dữ liệu: tổng hợp, PivotTable, xu hướng, so sánh, tô điểm bất thường, biểu đồ, nhận định. Dùng khi phân tích, thống kê, báo cáo nhanh.",
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
    brief: "Dựng slide báo cáo tiến độ dự án BĐS: pháp lý, thi công, tài chính, vướng mắc, kế hoạch tiếp. Dùng khi làm hoặc cập nhật báo cáo tiến độ.",
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
    brief: "Chuyển tài liệu (tệp đính kèm, Word hoặc Excel đang mở) thành bài trình chiếu có cấu trúc. Dùng khi làm slide từ tài liệu, báo cáo, văn bản.",
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
    brief: "Soát bài trình chiếu: chữ mẫu sót, chính tả, tiêu đề lệch, quá nhiều chữ, chữ nhỏ, số liệu lệch. Dùng khi kiểm tra, soát, hoàn thiện slide.",
    description:
      "Soát bài trình chiếu PowerPoint: chữ mẫu còn sót, lỗi chính tả, tiêu đề lệch vị trí/kiểu, slide quá nhiều chữ, cỡ chữ quá nhỏ, thứ tự và mạch nội dung, số liệu không khớp; sửa khi được đồng ý. Dùng khi người dùng muốn kiểm tra, soát, hoàn thiện bài trình chiếu trước khi gửi.",
    hosts: ["PowerPoint"],
    source: "builtin",
    enabled: true,
    triggers: ["kiem tra bai trinh chieu", "soat slide", "ra soat slide", "hoan thien slide", "deck check", "kiem tra slide"],
    body: `# Kiểm tra bài trình chiếu

1. ppt_check_deck trước (mã tự dò: phông lệch mẫu, chữ quá nhỏ, slide thiếu tiêu đề, chữ ước tràn khung, hình ra ngoài slide, khung chồng nhau), rồi ppt_get_overview toàn bài. Slide bị nghi tràn hoặc chồng: ppt_get_slide_image để nhìn tận mắt.
2. Soát thêm: chữ mẫu còn sót ("Click to add…", "Lorem", "[cần bổ sung]"); lỗi chính tả; tiêu đề có dấu chấm cuối hoặc khác vị trí; slide > 6 ý hoặc chữ dài; số liệu mâu thuẫn giữa các slide; thứ tự kể chuyện; slide thiếu yếu tố trực quan.
3. Báo bảng: Slide số | Vấn đề | Đề xuất.
4. Sửa khi được đồng ý: ppt_set_shape_text, ppt_format_text (phông, cỡ chữ theo mẫu), ppt_update_shape, ppt_set_layout (slide thiếu ô tiêu đề), ppt_move_slide; có Hoàn tác. Sửa xong slide nào thì chụp lại slide đó một lần để xác nhận.`,
  },
  {
    name: "kiem-tra-tuan-thu",
    brief: "Kiểm tra một việc định làm có đúng quy định, cần phê duyệt gì. Dùng khi hỏi 'có đúng quy định không', 'cần xin phép gì'.",
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
    brief: "Soát văn bản trước khi trình ký, kết luận ĐẠT / CÓ VẤN ĐỀ. Dùng khi nói 'trước khi ký', 'trình ký', 'bản cuối'.",
    description:
      "Soát hợp đồng hoặc văn bản trong Word trước khi trình ký: còn sửa đổi chưa duyệt, nhận xét chưa xử lý, phụ lục được nhắc tới có đủ không, tên pháp nhân, người ký, chức danh, ủy quyền, ngày tháng, số tiền bằng số và bằng chữ, đánh số và tham chiếu, chữ mẫu còn sót. Kết luận ĐẠT hoặc CÓ VẤN ĐỀ kèm danh sách cần sửa. Dùng khi người dùng nói 'trước khi ký', 'trình ký', 'chốt bản ký', 'kiểm tra bản cuối'.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["truoc khi ky", "trinh ky", "ban ky", "chot ban", "ban cuoi", "san sang ky", "ky ket", "signature"],
    body: `# Kiểm tra trước khi ký (phỏng theo plugin legal/signature-request, phần kiểm tra trước ký)

1. word_list_revisions: còn sửa đổi chưa chấp nhận hoặc từ chối → CÓ VẤN ĐỀ, vì bản ký phải là bản sạch. Không tự chấp nhận tất cả; hỏi người dùng.
2. word_list_comments: còn nhận xét chưa xử lý → liệt kê.
3. Xem mục lục điều khoản (word_get_outline) để soát đánh số, rồi đọc toàn văn (word_read_paragraphs, gọi tiếp theo next_start) và kiểm:
   - Tên pháp nhân, mã số doanh nghiệp, địa chỉ, người đại diện, chức danh giống nhau ở phần chủ thể, phần ký và mọi chỗ nhắc lại.
   - Người ký có thẩm quyền: là người đại diện theo pháp luật, hoặc có văn bản ủy quyền được dẫn chiếu (số, ngày).
   - Phụ lục, biểu mẫu, bản vẽ được nhắc tới đều có trong văn bản hoặc ghi rõ là kèm theo; số thứ tự phụ lục liên tục.
   - Ngày tháng: ngày ký để trống hoặc đúng; thời hạn và mốc tiến độ không mâu thuẫn.
   - Số tiền: bằng số khớp bằng chữ; đơn vị tiền tệ; đã gồm hay chưa gồm thuế ghi nhất quán.
   - Đánh số điều, khoản; tham chiếu chéo trỏ đúng; thuật ngữ định nghĩa dùng nhất quán.
   - Số bản, ngôn ngữ ưu tiên, hiệu lực; ký nháy từng trang nếu áp dụng.
   - Chữ mẫu còn sót: "…", "[ ]", "xxx", "[cần bổ sung]", chỗ tô vàng.
4. Kết quả: "Kiểm tra trước ký: ĐẠT" hoặc "CÓ VẤN ĐỀ", kèm bảng Hạng mục | Kết quả | Vị trí (số điều, [¶n]) | Cần làm. Hỏi người dùng có muốn đặt nhận xét tại từng chỗ (word_add_comment) không.
5. Add-in không gửi văn bản đi ký điện tử; nhắc lưu bản đã ký vào hồ sơ.`,
  },
  {
    name: "chuan-bi-hop",
    brief: "Tổng hợp tài liệu chuẩn bị họp, đàm phán có nội dung pháp lý. Dùng khi nói 'chuẩn bị họp', 'tài liệu họp', 'mai đi đàm phán'.",
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
    brief: "Soạn văn bản trả lời yêu cầu, câu hỏi, khiếu nại pháp lý; nhận diện ca phải chuyển cấp trên. Dùng khi nói 'soạn văn bản trả lời', 'phản hồi khách hàng'.",
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
5. Văn bản gửi cơ quan nhà nước hoặc cần đúng thể thức văn bản hành chính (quốc hiệu, số và ký hiệu, trích yếu, nơi nhận): đọc kỹ năng soan-thao-van-ban-hanh-chinh (use_skill) và dựng bằng word_create_admin_document.
6. Văn bản khác (thư trả lời khách hàng, đối tác): chèn vào Word bằng word_insert_paragraphs (Theo dõi thay đổi); báo các chỗ cần điền [ ].`,
  },
  {
    name: "soan-thao-van-ban-hanh-chinh",
    brief: "Soạn, sửa, soát thể thức văn bản hành chính (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản, kế hoạch). Dùng khi nói 'soạn công văn', 'tờ trình', 'đúng thể thức chưa'.",
    description:
      "Soạn, sửa và soát thể thức văn bản hành chính trong Word: công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản, kế hoạch. Dựng trọn văn bản đúng thể thức (quốc hiệu, tiêu ngữ, tên cơ quan, số và ký hiệu, địa danh ngày tháng, trích yếu, kính gửi, phần ký, nơi nhận, lề, phông, cỡ chữ) và kiểm lại. Dùng khi người dùng nói 'soạn công văn', 'soạn tờ trình', 'viết báo cáo gửi…', 'soạn quyết định', 'thông báo', 'giấy mời', 'biên bản', 'kiểm tra thể thức', 'đúng thể thức chưa'.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: [
      "cong van", "to trinh", "van ban hanh chinh", "dung the thuc", "sai the thuc", "the thuc van ban", "kiem tra the thuc", "soat the thuc", "soan quyet dinh", "du thao quyet dinh",
      "soan thong bao", "viet thong bao", "du thao thong bao", "giay moi", "soan bien ban", "lap bien ban", "viet bien ban", "soan bao cao", "du thao bao cao", "viet bao cao",
      "soan ke hoach", "lap ke hoach", "soan van ban gui", "soan van ban de nghi", "quoc hieu",
    ],
    body: `# Soạn thảo văn bản hành chính

Áp dụng chuẩn thể thức và kỹ thuật trình bày văn bản hành chính đang có hiệu lực (bắt buộc với cơ quan, tổ chức nhà nước và doanh nghiệp nhà nước; doanh nghiệp khác nên theo, nhất là văn bản gửi cơ quan nhà nước). Không dùng cho văn bản quy phạm pháp luật. Mẫu do bộ, ngành ban hành riêng (đơn, văn bản đề nghị theo mẫu thủ tục): giữ nguyên bố cục và câu chữ của mẫu, chỉ chỉnh phông, cỡ chữ, lề.

## A. Soạn mới
1. Xác định trước khi viết: loại văn bản; cơ quan ban hành (và cơ quan chủ quản nếu có); gửi ai; nội dung chính; ai ký và ký với quyền hạn gì. Người dùng đã nêu đủ ý chính thì soạn ngay; chỉ hỏi khi thiếu tên cơ quan ban hành hoặc thiếu nội dung chính.
2. KHÔNG bịa số văn bản, ngày ký, tên người ký, ký hiệu đơn vị, số liệu, căn cứ pháp lý. Chưa biết thì bỏ trống tham số: công cụ tự chừa chỗ. Số văn bản luôn để trống cho văn thư cấp.
3. Dựng bằng MỘT lần gọi word_create_admin_document (công cụ tự trình bày đúng mọi thành phần, đặt khổ giấy, lề, số trang khi văn bản đang trống). Không tự chèn từng dòng quốc hiệu, bảng, phần ký bằng công cụ khác.
   - doc_type: cong_van, to_trinh, bao_cao, quyet_dinh, thong_bao, giay_moi, bien_ban, ke_hoach, khac (kèm title). Hợp đồng (kể cả khi người dùng muốn "theo thể thức hành chính"): doc_type hop_dong, làm theo kỹ năng soan-hop-dong (use_skill), không dùng khac.
   - symbol: chữ viết tắt tên loại + chữ viết tắt tên cơ quan, ví dụ TTr-ABC, BC-ABC, QĐ-ABC, TB-ABC, GM-ABC, BB-ABC, KH-ABC. Công văn không có chữ viết tắt tên loại: chữ viết tắt cơ quan + đơn vị soạn thảo, ví dụ ABC-PC. Chỉ điền khi người dùng đã cho chữ viết tắt của cơ quan hoặc tài liệu có sẵn; không biết thì bỏ trống, không tự đặt.
   - subject: trích yếu ngắn, nêu đúng việc (công văn: phần sau "V/v").
   - to: nơi nhận ở phần Kính gửi (công văn, tờ trình, báo cáo gửi cấp trên). recipients: các dòng Nơi nhận.
   - body: mỗi dòng một đoạn; "# " đề mục canh giữa; "## " đề mục đậm; cả dòng trong dấu gạch dưới _…_ là chữ nghiêng (dòng căn cứ). Body chỉ là phần nội dung: không chép quốc hiệu, tên cơ quan, tên loại, số, ngày, phần ký, nơi nhận vào body (công cụ đã dựng).
   - signer_authority (TM., KT., TL., TUQ., Q. nếu có), signer_title, signer_name (không ghi học hàm, học vị).
4. Viết nội dung theo loại:
   - Công văn: mở đầu nêu lý do hoặc dẫn văn bản liên quan; nội dung trình bày từng ý, đề nghị cụ thể, thời hạn; câu kết ("Trân trọng." hoặc đề nghị phối hợp). Văn bản gửi cấp trên dùng "kính đề nghị", gửi ngang cấp hoặc cấp dưới dùng "đề nghị".
   - Tờ trình: "## I. …" sự cần thiết / căn cứ; "## II. …" nội dung trình (phương án, số liệu, kinh phí, tiến độ); "## III. …" kiến nghị, đề xuất cụ thể để người có thẩm quyền quyết; kết "Kính trình … xem xét, quyết định."
   - Báo cáo: tình hình và kết quả (có số liệu, mốc thời gian); tồn tại, nguyên nhân; phương hướng; đề xuất, kiến nghị.
   - Quyết định: dòng "# <THẨM QUYỀN BAN HÀNH>" (ví dụ "# GIÁM ĐỐC CÔNG TY …"); các dòng căn cứ nghiêng "_Căn cứ …;_", dòng cuối "_Theo đề nghị của …._" kết bằng dấu chấm; "# QUYẾT ĐỊNH:"; rồi "## Điều 1. …", "## Điều 2. …"; điều cuối nêu hiệu lực và người chịu trách nhiệm thi hành. Nơi nhận ghi "Như Điều …".
   - Thông báo: nội dung thông báo, đối tượng, thời gian, yêu cầu thực hiện, đầu mối liên hệ.
   - Giấy mời: dòng "<cơ quan> trân trọng kính mời: …"; "Tới dự: …"; "Chủ trì: …"; "Thời gian: …"; "Địa điểm: …".
   - Biên bản: thời gian bắt đầu, địa điểm, thành phần tham dự, chủ trì, thư ký, nội dung, giờ kết thúc; ký hai bên: left_signer_title "THƯ KÝ", signer_title "CHỦ TỌA".
   - Đề mục trong nội dung: phần "I., II.", mục "1., 2.", điểm "a), b)"; đánh số liên tục, không nhảy, không trùng.
5. Căn cứ pháp lý chỉ ghi khi có trong tài liệu, tệp đính kèm hoặc người dùng cung cấp; lần đầu viện dẫn ghi đủ tên loại, số, ký hiệu, ngày, cơ quan ban hành, trích yếu. Không có thì ghi "[cần bổ sung căn cứ]" và báo người dùng.
6. Viết hoa: tên cơ quan viết hoa chữ đầu các từ chỉ loại hình, chức năng (Sở Xây dựng, Ủy ban nhân dân tỉnh …); "khoản", "điểm" viết thường, "Điều", "Chương" viết hoa khi viện dẫn; "cơ quan nhà nước", "quản lý nhà nước" viết thường chữ "nhà nước"; chức vụ viết hoa khi đi liền tên người.
7. Sau khi dựng: gọi word_check_page_format để xác nhận lề, khổ giấy, phông, cỡ chữ; báo người dùng danh sách chỗ còn trống (trường blanks của công cụ) cần điền trước khi ký. Sửa câu chữ sau đó thì dùng word_set_paragraph_text trên đúng đoạn, không dựng lại cả văn bản.

## B. Soát và sửa văn bản có sẵn
1. Đọc văn bản (word_read_paragraphs), kiểm trang bằng word_check_page_format, đọc đầu và chân trang bằng word_read_parts.
2. Đối chiếu với chuẩn dưới đây, lập bảng: Thành phần | Chuẩn | Hiện trạng | Xử lý.
3. Người dùng chỉ nhờ KIỂM thì chỉ báo kết quả theo bảng, không sửa. Khi được nhờ sửa thể thức: phông, cỡ, kiểu chữ, canh lề, thụt đầu dòng, cách đoạn, giãn dòng (word_format_text với font_name, font_size, bold, italic, alignment, first_line_indent_mm, space_after_pt, line_spacing), viết hoa chỗ rõ ràng, dấu câu của Kính gửi và Nơi nhận; lề và khổ giấy thì nêu số đo hiện tại và số đo sẽ đặt rồi mới gọi word_set_page_format (thay đổi này áp cho cả văn bản). Hợp đồng và văn bản không phải văn bản hành chính: không áp chuẩn này trừ khi người dùng yêu cầu.
4. KHÔNG tự sửa nội dung: câu chữ, số liệu, ngày tháng, số hiệu văn bản viện dẫn, người ký. Thấy sai thì ghi nhận xét (word_add_comment) để người dùng quyết.
5. Phần đầu và phần ký là bảng không viền hai cột (đoạn có trường cell khi đọc): tên cơ quan, số ở ô trái nằm NGANG hàng quốc hiệu, ngày ở ô phải — đúng thể thức, không đổi thành trên dưới. Bớt hay sửa một dòng ở đó: chỉ xóa hoặc sửa đúng đoạn đó; không xóa rồi chèn lại để đổi chỗ.

## C. Chuẩn thể thức (số đo)
- Giấy A4 đứng. Lề trên 20–25 mm, dưới 20–25 mm, trái 30–35 mm, phải 15–20 mm. Phông Times New Roman, màu đen.
- Nội dung: cỡ 13–14, đứng, canh đều hai lề, thụt đầu dòng 10–12,7 mm (không dùng tab hay dấu cách), cách đoạn tối thiểu 6 pt, giãn dòng từ đơn đến 1,5; kết thúc bằng "./.".
- Quốc hiệu: in hoa, đậm, cỡ 12–13. Tiêu ngữ: chữ thường, đậm, cỡ 13–14, có đường kẻ liền bên dưới dài bằng dòng chữ. Hai cỡ tăng giảm cùng nhau (12 và 13, hoặc 13 và 14).
- Tên cơ quan chủ quản: in hoa, đứng, cỡ 12–13. Tên cơ quan ban hành: in hoa, đậm, cỡ 12–13, có đường kẻ bên dưới dài khoảng một phần ba đến một nửa dòng chữ.
- Số và ký hiệu: "Số: 15/QĐ-ABC", cỡ 13, đặt dưới tên cơ quan. Địa danh và ngày tháng: nghiêng, cỡ 13–14, cùng hàng với số, dưới tiêu ngữ; ngày dưới 10 và tháng 1, 2 thêm số 0 ở trước.
- Tên loại: in hoa, đậm, cỡ 13–14, canh giữa. Trích yếu: chữ thường, đậm, ngay dưới tên loại, có đường kẻ bên dưới. Công văn: "V/v …" cỡ 12–13, đứng, dưới số và ký hiệu.
- Kính gửi: cỡ 13–14; nhiều nơi nhận thì mỗi nơi một dòng có gạch đầu dòng, cuối dòng dấu chấm phẩy, dòng cuối dấu chấm.
- Phần ký: quyền hạn và chức vụ in hoa, đậm; chừa bốn đến năm dòng cho chữ ký và dấu; họ tên chữ thường, đậm.
- Nơi nhận: chữ "Nơi nhận:" cỡ 12, nghiêng, đậm; các dòng bên dưới cỡ 11, đứng; dòng cuối "- Lưu: VT, <đơn vị soạn thảo>."
- Số trang: chữ số Ả Rập, cỡ 13–14, canh giữa ở lề trên, không hiện ở trang đầu.`,
  },
  {
    name: "soan-hop-dong",
    brief: "Soạn hợp đồng (xây dựng, thi công, tư vấn, mua bán) đúng thể thức, ký hai bên. Dùng khi nói 'soạn hợp đồng', 'mẫu hợp đồng'.",
    description:
      "Soạn mới hợp đồng trong Word đúng thể thức trình bày: quốc hiệu, tiêu ngữ (tên đơn vị bên trái nếu cần), tên và số hợp đồng, căn cứ, thông tin các bên, các điều khoản, phần ký hai bên; sửa phần đầu, phần ký mà không làm lệch bố cục. Dùng khi người dùng nói 'soạn hợp đồng', 'dự thảo hợp đồng', 'soạn mẫu hợp đồng xây dựng', 'lập hợp đồng thi công', 'viết hợp đồng'.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: [
      "soan hop dong", "soan thao hop dong", "du thao hop dong", "lap hop dong", "viet hop dong", "tao hop dong",
      "soan mau hop dong", "tao mau hop dong", "lap mau hop dong", "du thao mau hop dong",
    ],
    body: `# Soạn hợp đồng

## A. Soạn mới
1. Nắm trước khi viết: loại hợp đồng (thi công xây dựng, tư vấn, cung cấp vật tư, mua bán, dịch vụ…); Bên A, Bên B; đối tượng (công trình, gói thầu, địa điểm, khối lượng); giá và cách tính giá; thời hạn; người đại diện ký. Soạn MẪU hợp đồng, hoặc người dùng đã nêu loại hợp đồng: soạn ngay, chỗ chưa có thông tin để "…", rồi kê các chỗ đó; không hỏi trước. Chỉ hỏi (một lần) khi người dùng muốn hợp đồng cho một việc cụ thể mà chưa rõ cả loại lẫn đối tượng.
2. KHÔNG bịa: số hợp đồng, ngày ký, mã số thuế, tài khoản, địa chỉ, người đại diện, giá trị, căn cứ pháp lý. Chưa biết thì để "…" cho người dùng điền; căn cứ chỉ ghi khi có trong tài liệu, tệp đính kèm hoặc người dùng đưa. KHÔNG tự viết số hiệu, ngày ban hành của luật, nghị định theo trí nhớ (rất dễ sai số, sai năm, có khi văn bản không tồn tại): chưa có nguồn thì để dòng "_Căn cứ [cần bổ sung căn cứ pháp lý];_" và báo người dùng.
3. Dựng bằng MỘT lần gọi word_create_admin_document với doc_type "hop_dong" (không dùng "khac"):
   - title: tên hợp đồng, ví dụ "HỢP ĐỒNG THI CÔNG XÂY DỰNG".
   - org: chỉ khi người dùng muốn tên đơn vị ở góc trái, ngang hàng với quốc hiệu bên phải (khi đó "Số: …" nằm dưới tên đơn vị). Không có org: quốc hiệu, tiêu ngữ canh giữa, "Số: …" dưới tên hợp đồng.
   - number (đủ dạng nếu người dùng cho), symbol (ví dụ HĐ-XD) chỉ khi người dùng cho; subject: dòng dưới tên hợp đồng nếu cần (ví dụ "Gói thầu: … thuộc dự án …").
   - body: CHỈ phần nội dung, bắt đầu từ các dòng căn cứ. KHÔNG chép vào body: quốc hiệu, tiêu ngữ, tên đơn vị, tên hợp đồng, dòng "Số:", địa danh ngày tháng, khối "ĐẠI DIỆN BÊN A/B" — công cụ đã tự dựng, chép lại là thừa dòng.
   - left_signer_title, left_signer_name: chức vụ, họ tên đại diện Bên A (cột trái); signer_title, signer_name: Bên B (cột phải). Nhãn "ĐẠI DIỆN BÊN A", "ĐẠI DIỆN BÊN B" tự có; muốn ghi khác thì left_signer_authority, signer_authority.
4. Body, mỗi dòng một đoạn:
   - Căn cứ: mỗi căn cứ một dòng nghiêng "_Căn cứ …;_", dòng căn cứ cuối kết bằng dấu chấm.
   - "Hôm nay, ngày … tháng … năm …, tại …, chúng tôi gồm:"
   - "## BÊN GIAO THẦU (BÊN A): <tên>" rồi từng dòng: "Đại diện: … Chức vụ: …", "Địa chỉ: …", "Điện thoại: …", "Mã số thuế: …", "Tài khoản số: … tại …". Bên B "## BÊN NHẬN THẦU (BÊN B): …" cùng các dòng đó.
   - "Hai bên thống nhất ký kết hợp đồng … với các điều khoản sau:"
   - Điều khoản: "## Điều 1. <tên điều>"; khoản "1.1.", "1.2." hoặc "1.", "2."; điểm "a)", "b)". Đánh số liên tục, không nhảy, không trùng.
5. Hợp đồng thi công xây dựng thường có các điều: đối tượng, khối lượng công việc; chất lượng, yêu cầu kỹ thuật, nghiệm thu, bàn giao; thời gian và tiến độ; giá hợp đồng, tạm ứng, thanh toán (hồ sơ, thời hạn, đồng tiền); điều chỉnh giá, khối lượng; bảo đảm thực hiện hợp đồng, bảo lãnh tạm ứng; bảo hành; quyền và nghĩa vụ của mỗi bên; an toàn lao động, vệ sinh môi trường, phòng chống cháy nổ; bảo hiểm; rủi ro, bất khả kháng; tạm dừng, chấm dứt; thưởng, phạt vi phạm, bồi thường; giải quyết tranh chấp; điều khoản chung (hiệu lực, số bản mỗi bên giữ). Số tiền ghi cả bằng số và bằng chữ, hai chỗ khớp nhau.
6. Xong: báo người dùng danh sách chỗ còn trống (blanks) và lưu ý (notes) của công cụ.

## B. Sửa phần đầu, phần ký
Phần đầu (tên đơn vị, "Số:" bên trái; quốc hiệu, tiêu ngữ bên phải) và phần ký là bảng không viền hai cột. Khi đọc (word_read_paragraphs), đoạn trong ô có trường cell: hai ô cùng hàng nằm NGANG nhau, chỉ số đoạn đi hết ô trái rồi mới sang ô phải. Tên đơn vị bên trái, quốc hiệu bên phải là đúng thể thức, không "sắp lại" thành trên dưới.
1. Thừa một dòng (ví dụ tên hợp đồng lặp hai lần): đọc lại, xóa ĐÚNG đoạn đó bằng word_delete_paragraphs (chỉ đoạn đó, expect là chữ của đoạn), không đụng đoạn khác, rồi đọc lại để xác nhận.
2. Đổi chữ một dòng: word_set_paragraph_text trên đúng đoạn đó. Định dạng: word_format_text.
3. Không xóa rồi chèn lại để đổi chỗ, không chèn đoạn mới vào phần đầu. Phần đầu đã lệch nhiều thì báo người dùng và đề nghị dựng lại trên văn bản trống, không tự vá từng dòng.
4. Người dùng chỉ nêu một lỗi thì chỉ sửa lỗi đó.`,
  },
  {
    name: "thiet-ke-bieu-do",
    brief: "Chọn dạng và trình bày biểu đồ đúng: theo việc người xem cần, màu chủ đề, một trục tung, tiêu đề nêu ý chính. Dùng khi vẽ, sửa biểu đồ, hỏi nên dùng biểu đồ gì.",
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
  {
    name: "dien-mau",
    brief: "Điền thông tin vào mẫu có sẵn (ô điền mẫu, chỗ trống '……', '[●]', ô bảng trống), thiếu thì để trống. Dùng khi nói 'điền mẫu', 'điền tờ khai', 'điền hợp đồng'.",
    description:
      "Điền thông tin vào mẫu Word có sẵn (hợp đồng mẫu, tờ khai, giấy đề nghị, biểu mẫu thủ tục): ô điền mẫu, chỗ trống '……', '____', '[●]', ô bảng trống. Lấy thông tin từ câu của người dùng, tệp đính kèm, bảng Excel đang mở; thiếu thì để trống và kê ra, không bịa. Dùng khi người dùng nói 'điền mẫu', 'điền vào mẫu', 'điền thông tin', 'điền hợp đồng', 'điền tờ khai', 'điền biểu mẫu', 'hoàn thiện mẫu'.",
    hosts: ["Word"],
    source: "builtin",
    enabled: true,
    triggers: ["dien mau", "dien mau hop dong", "dien vao", "dien giup", "dien ho", "dien cac cho", "dien vao mau", "dien thong tin", "dien hop dong", "dien to khai", "dien bieu mau", "dien cho trong", "dien vao cho trong", "hoan thien mau", "dien giay", "dien don"],
    body: `# Điền mẫu

1. Gọi word_list_form_fields một lần để biết mọi chỗ cần điền (mã, nhãn, loại ô, danh sách mục của ô chọn, định dạng ngày). Kết quả có next_start là chưa hết: gọi tiếp.
2. Gom thông tin trước khi điền:
   - từ câu của người dùng;
   - từ tệp đính kèm: PDF đọc theo trang bằng read_attachment_pages; tệp .docx đã có chữ trong tin nhắn;
   - từ Excel đang mở: read_other_app (nếu đã kết nối).
3. Ghép từng chỗ với thông tin theo NHÃN đứng trước chỗ đó (ví dụ "Số CCCD:", "cấp ngày", "tại"), không theo thứ tự xuất hiện. Một chỗ trống mà nhãn không rõ (nhãn rỗng hoặc chỉ là dấu câu) thì xem chữ đứng sau (after) và đoạn chứa nó.
4. Điền TẤT CẢ trong MỘT lần gọi word_fill_form, mỗi chỗ kèm label đúng như danh sách. Không điền từng chỗ một.
5. KHÔNG BỊA: số giấy tờ (CCCD, mã số thuế, số thửa, số giấy chứng nhận), ngày tháng, số tiền, địa chỉ, tên người — không có trong thông tin được đưa thì bỏ chỗ đó, không đoán, không lấy ví dụ.
6. Cách viết:
   - Ô đánh dấu: true hoặc false. Ô chọn: đúng một mục trong danh sách của ô; thông tin không khớp mục nào thì không điền và nêu lại cho người dùng.
   - Ngày: dd/MM/yyyy (công cụ tự viết theo định dạng của ô). Chỗ trống "ngày   tháng   năm" thay cả cụm, ví dụ "ngày 18 tháng 11 năm".
   - Số tiền: điền cả bằng số và bằng chữ, hai chỗ phải khớp nhau; số viết kiểu Việt Nam (1.250.000.000). Tên riêng giữ đúng chữ hoa như nguồn.
   - Chỗ trống chỉ là chỗ để ký, đóng dấu: không điền.
7. Kết quả của word_fill_form đã kèm amount_check (soát số tiền bằng số và bằng chữ); không cần gọi thêm công cụ soát hay đọc lại văn bản để kiểm. Có chỗ lệch: nếu do bạn viết chữ sai thì sửa đúng cụm đó bằng word_replace_text (mã chỗ điền cũ không dùng được nữa sau khi điền); nếu chính thông tin người dùng đưa có số và chữ lệch nhau thì KHÔNG tự chọn — hỏi lại người dùng.
8. Trả lời ngắn: (a) đã điền bao nhiêu chỗ, (b) chỗ bỏ qua và lý do công cụ trả, (c) chỗ còn trống vì thiếu thông tin — kê theo nhãn để người dùng bổ sung.`,
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
