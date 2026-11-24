/** Bản 4.11.10: sửa theo buổi thử model miễn phí soạn hợp đồng và văn bản hành chính (23–24/11/2026,
 *  `tasks/thu-mien-phi-hdcn.md`, `tasks/thu-mien-phi-vbhc.md`). Mỗi ca lấy đúng dạng tham số model đã gửi. */
import * as assert from "assert";
import { ParaBlock, TableBlock, adminDocBrief, bodyLinesOf, normalizeSigner, planAdminDocument } from "../../src/core/adminDoc";
import { toolFailureClaim } from "../../src/core/claims";
import { BUILTIN_SKILLS } from "../../src/core/skillpack";
import { WORD_TOOL_GROUPS, requestGroups } from "../../src/core/toolgroups";
import { isVietSyllable, strangeWords, weekdayIssues } from "../../src/core/textcheck";

const paras = (p: ReturnType<typeof planAdminDocument>) => p.blocks.filter((b) => b.kind === "para") as ParaBlock[];
const tables = (p: ReturnType<typeof planAdminDocument>) => p.blocks.filter((b) => b.kind === "table") as TableBlock[];
const allText = (p: ReturnType<typeof planAdminDocument>) =>
  p.blocks.flatMap((b) => (b.kind === "table" ? b.rows.flatMap(([l, r]) => [...l, ...r]).map((x) => x.text) : [b.text]));
const ORG = "Công ty Cổ phần Đầu tư Anh Quân";

describe("4.11.10 · Nội dung model miễn phí đưa vào công cụ dựng văn bản", () => {
  it("body là chuỗi MẢNG JSON (Nemotron 3 Ultra): tách thành từng đoạn, không dồn một đoạn", () => {
    const body = JSON.stringify(["_Căn cứ Luật Doanh nghiệp năm 2020;_", "# QUYẾT NGHỊ:", "## Điều 1. Thông qua việc vay vốn.", "## Điều 2. Nghị quyết có hiệu lực kể từ ngày ký."]);
    assert.deepStrictEqual(bodyLinesOf(body).length, 4);
    const p = planAdminDocument({ doc_type: "nghi_quyet", org: ORG, body, signer_authority: "TM. HỘI ĐỒNG QUẢN TRỊ", signer_title: "CHỦ TỊCH" });
    const t = paras(p).map((x) => x.text);
    assert.ok(t.includes("QUYẾT NGHỊ:"));
    assert.ok(!t.some((x) => x.startsWith("[")), t.join(" | "));
  });
  it("chuỗi chỉ có '\\n' viết bằng chữ cũng tách dòng; gạch dưới đầu dòng thành gạch ngang", () => {
    assert.deepStrictEqual(bodyLinesOf("Dòng một\\nDòng hai"), ["Dòng một", "Dòng hai"]);
    assert.deepStrictEqual(bodyLinesOf("_ Tên khoản vay: 150 tỷ đồng.\n_Căn cứ Điều lệ;_"), ["- Tên khoản vay: 150 tỷ đồng.", "_Căn cứ Điều lệ;_"]);
  });
  it("thẻ trích dẫn Cohere <co>…</co: 2:[0]> bị bỏ khỏi nội dung và tham số", () => {
    const p = planAdminDocument({ doc_type: "thong_bao", org: ORG, subject: "<co>Lịch nghỉ Tết</co: 0:[1]>", body: "## Điều 1. <co>Thông qua việc vay vốn</co: 2:[0]>" });
    assert.ok(!allText(p).some((x) => /<\/?co/.test(x)), allText(p).join(" | "));
  });
  it("'## Điều 1.' kèm lời văn dài: chỉ 'Điều 1.' đậm; điều có tên ngắn vẫn đậm cả dòng", () => {
    const p = planAdminDocument({
      doc_type: "quyet_dinh",
      org: ORG,
      body: "# QUYẾT ĐỊNH:\n## Điều 1. Bổ nhiệm ông Trần Văn Bình giữ chức Phó Tổng Giám đốc kể từ ngày 01 tháng 12 năm 2026, thời hạn 03 năm.\n## Điều 2. Hiệu lực thi hành",
    });
    const d1 = paras(p).find((x) => x.text.startsWith("Điều 1."))!;
    assert.strictEqual(d1.bold, undefined);
    assert.strictEqual(d1.boldLead, "Điều 1.");
    assert.strictEqual(paras(p).find((x) => x.text.startsWith("Điều 2."))!.bold, true);
  });
});

describe("4.11.10 · Thể thức theo loại văn bản", () => {
  it("dòng 'Kính gửi' chép lại đầu body bị bỏ, giá trị dùng làm Kính gửi", () => {
    const p = planAdminDocument({ doc_type: "cong_van", org: ORG, subject: "Đề nghị", body: "Kính gửi: Sở Xây dựng tỉnh Bình Minh,\nCông ty kính đề nghị Sở xem xét.\nTrân trọng," });
    const t = paras(p).map((x) => x.text);
    assert.strictEqual(t.filter((x) => x.startsWith("Kính gửi")).length, 1);
    assert.ok(t.includes("Kính gửi: Sở Xây dựng tỉnh Bình Minh"), t.join(" | "));
    // "Trân trọng," → "Trân trọng./."
    assert.ok(t.includes("Trân trọng./."), t.join(" | "));
  });
  it("quyết định không có Kính gửi: nơi nhận chuyển sang Nơi nhận", () => {
    const p = planAdminDocument({ doc_type: "quyet_dinh", org: ORG, to: ["Các thành viên Hội đồng quản trị", "Ban Giám đốc"], body: "## Điều 1. Bổ nhiệm …" });
    const t = allText(p);
    assert.ok(!t.some((x) => x.startsWith("Kính gửi")));
    assert.ok(t.includes("- Các thành viên Hội đồng quản trị;"), t.join(" | "));
  });
  it("Nơi nhận 'Như Điều 3.' không thành 'Như Điều 3.;'", () => {
    const p = planAdminDocument({ doc_type: "quyet_dinh", org: ORG, recipients: ["Như Điều 3.", "Lưu: VT"], body: "## Điều 1. …" });
    assert.ok(allText(p).includes("- Như Điều 3;"), allText(p).join(" | "));
  });
  it("nghị quyết: loại riêng, nhận cả khac/quyet_dinh kèm tên NGHỊ QUYẾT; 'QUYẾT ĐỊNH:' đổi thành 'QUYẾT NGHỊ:'", () => {
    for (const doc_type of ["nghi_quyet", "khac", "quyet_dinh"]) {
      const p = planAdminDocument({ doc_type, title: "NGHỊ QUYẾT VỀ VIỆC PHÊ DUYỆT CÁC NỘI DUNG", org: ORG, body: "# QUYẾT ĐỊNH:\n## Điều 1. Thông qua …" });
      assert.strictEqual(p.type, "nghi_quyet", doc_type);
      const t = paras(p).map((x) => x.text);
      assert.ok(t.includes("NGHỊ QUYẾT") && t.includes("Về việc phê duyệt các nội dung"), t.join(" | "));
      assert.ok(t.includes("QUYẾT NGHỊ:") && !t.includes("QUYẾT ĐỊNH:"));
    }
  });
  it("số đủ dạng có năm ('15/2026/NQ-HĐQT') giữ nguyên", () => {
    const p = planAdminDocument({ doc_type: "nghi_quyet", org: ORG, number: "15/2026/NQ-HĐQT", body: "## Điều 1. …" });
    assert.ok(allText(p).includes("Số: 15/2026/NQ-HĐQT"));
    assert.ok(!p.blanks.includes("ký hiệu văn bản"));
  });
  it("tờ trình có dòng 'QUYẾT ĐỊNH:' (Cohere): báo trong checks", () => {
    const p = planAdminDocument({ doc_type: "to_trinh", org: ORG, to: ["Hội đồng quản trị"], body: "# QUYẾT ĐỊNH:\n## I. Sự cần thiết" });
    assert.ok(p.checks?.some((c) => c.includes("QUYẾT ĐỊNH")), String(p.checks));
  });
  it("dòng tên loại in hoa kèm trích yếu ở đầu body (biên bản): bỏ, phần sau thành trích yếu", () => {
    const p = planAdminDocument({ doc_type: "bien_ban", org: ORG, body: "BIÊN BẢN ĐẠI HỘI ĐỒNG CỔ ĐÔNG BẤT THƯỜNG NĂM 2026\n## I. Thời gian, địa điểm" });
    const t = paras(p).map((x) => x.text);
    assert.ok(t.includes("Đại hội đồng cổ đông bất thường năm 2026"), t.join(" | "));
    assert.strictEqual(t.filter((x) => x.startsWith("BIÊN BẢN")).length, 1);
  });
});

describe("4.11.10 · Quyền hạn ký", () => {
  it("viết đủ chữ đổi ra dạng tắt; sai dạng bỏ; trùng chức vụ bỏ", () => {
    assert.strictEqual(normalizeSigner("KÝ THAY TỔNG GIÁM ĐỐC", "PHÓ TỔNG GIÁM ĐỐC").authority, "KT. TỔNG GIÁM ĐỐC");
    assert.strictEqual(normalizeSigner("Thay mặt Hội đồng quản trị", "Chủ tịch").authority, "TM. HỘI ĐỒNG QUẢN TRỊ");
    for (const bad of ["TGĐ", "ĐẠI DIỆN PHÁP LUẬT", "TM."]) {
      const r = normalizeSigner(bad, "TỔNG GIÁM ĐỐC");
      assert.strictEqual(r.authority, "", bad);
      assert.strictEqual(r.title, "TỔNG GIÁM ĐỐC");
    }
    const same = normalizeSigner("CHỦ TỊCH HỘI ĐỒNG QUẢN TRỊ", "CHỦ TỊCH HỘI ĐỒNG QUẢN TRỊ");
    assert.deepStrictEqual([same.authority, same.title], ["", "CHỦ TỊCH HỘI ĐỒNG QUẢN TRỊ"]);
    // Chức vụ đặt nhầm ở dòng quyền hạn, dòng chức vụ trống: đưa xuống
    assert.deepStrictEqual([normalizeSigner("TỔNG GIÁM ĐỐC", "").authority, normalizeSigner("TỔNG GIÁM ĐỐC", "").title], ["", "TỔNG GIÁM ĐỐC"]);
  });
  it("'KT. PHÓ …' bị báo; nghị quyết thiếu TM. bị báo", () => {
    assert.ok(normalizeSigner("KT. PHÓ TỔNG GIÁM ĐỐC", "PHÓ TỔNG GIÁM ĐỐC").checks.length === 1);
    const p = planAdminDocument({ doc_type: "nghi_quyet", org: ORG, body: "## Điều 1. …", signer_title: "CHỦ TỊCH HỘI ĐỒNG QUẢN TRỊ" });
    assert.ok(p.checks?.some((c) => c.includes("TM.")));
  });
  it("hợp đồng: chức vụ ghi lại nhãn 'ĐẠI DIỆN BÊN …' (lặp hoặc lộn bên) thì bỏ", () => {
    const p = planAdminDocument({
      doc_type: "hop_dong",
      title: "HỢP ĐỒNG CHUYỂN NHƯỢNG",
      body: "Điều 1. Đối tượng",
      left_signer_title: "ĐẠI DIỆN BÊN B",
      signer_title: "ĐẠI DIỆN BÊN A",
    });
    const sign = tables(p).pop()!.rows[0];
    const left = sign[0].map((x) => x.text).filter(Boolean);
    const right = sign[1].map((x) => x.text).filter(Boolean);
    assert.deepStrictEqual(left.slice(0, 1), ["ĐẠI DIỆN BÊN A"]);
    assert.ok(!left.includes("ĐẠI DIỆN BÊN B") && !right.includes("ĐẠI DIỆN BÊN A"), `${left} / ${right}`);
  });
  it("hợp đồng dán từ chat: phần ký tự viết và khối 'Lưu ý' ở cuối bị bỏ", () => {
    const p = planAdminDocument({
      doc_type: "hop_dong",
      title: "HỢP ĐỒNG CHUYỂN NHƯỢNG",
      body: ["Điều 7: Hiệu lực", "- Hợp đồng có hiệu lực từ ngày ký.", "---", "Bên chuyển nhượng", "(Ký, ghi rõ họ tên)", "Bên nhận chuyển nhượng", "(Ký, ghi rõ họ tên)", "Lưu ý: Cần đính kèm:", "- Giấy chứng nhận quyền sử dụng đất."].join("\n"),
    });
    const t = paras(p).map((x) => x.text);
    assert.ok(!t.some((x) => /^(Bên chuyển nhượng|Lưu ý|---|\(Ký)/.test(x)), t.join(" | "));
    assert.ok(p.notes.some((n) => n.includes("Lưu ý")));
  });
});

describe("4.11.10 · Kiểm nội dung (textcheck)", () => {
  it("thứ trong tuần lệch ngày bị báo, đúng thì không", () => {
    const bad = weekdayIssues(["- Ngày 01 tháng 01 năm 2027 (Thứ Ba): nghỉ Tết.", "- Đi làm lại từ thứ Hai, ngày 07 tháng 01 năm 2027."]);
    assert.strictEqual(bad.length, 2, bad.join(" | "));
    assert.ok(bad[0].includes("Thứ Sáu"));
    assert.deepStrictEqual(weekdayIssues(["Nghỉ ngày 01/01/2027 (thứ Sáu); đi làm lại từ ngày 04/01/2027 (thứ Hai)", "Chủ nhật, ngày 03/01/2027"]), []);
  });
  it("từ Latinh không phải tiếng Việt và chữ dính bị báo; tên riêng, chữ viết tắt, từ thường gặp thì không", () => {
    const odd = strangeWords(["không bị burden khác", "tranh chấp, cấmélé, tạm giữ hoặc competencia", "Bên A là chủthể có quyền", "đạt quórum theo luật"]);
    for (const w of ["burden", "cấmélé", "competencia", "chủthể", "quórum"]) assert.ok(odd.some((x) => x.startsWith(w)), `${w}: ${odd}`);
    assert.deepStrictEqual(strangeWords(["Dự án Khu nhà ở Anh Quân Riverside, Ngân hàng BIDV, gửi email cho khách; nghiêng, khuya, quốc, giữ gìn, thuở, ngoằn ngoèo"]), []);
    for (const w of ["nghiêng", "khuyến", "quyết", "người", "thuở", "giặt"]) assert.ok(isVietSyllable(w), w);
  });
  it("kết quả dựng văn bản có checks khi nội dung sai thứ hoặc lẫn từ lạ", () => {
    const p = planAdminDocument({ doc_type: "thong_bao", org: ORG, body: "Ngày 01/01/2027 (Thứ Năm): nghỉ Tết Dương lịch.\nĐọc số đồng hồ, meter điện." });
    assert.strictEqual(p.checks?.length, 2, String(p.checks));
  });
});

describe("4.11.10 · Cảnh báo, từ khóa, kỹ năng", () => {
  it("AI nói công cụ lỗi mà không gọi công cụ nào: cảnh báo; có gọi công cụ thì không", () => {
    const answer = "Công cụ soạn thảo vẫn đang lỗi nên bản này tôi ghi trực tiếp trong câu trả lời.";
    assert.ok(toolFailureClaim({ userText: "Soạn hợp đồng", answer, wrote: false, toolsCalled: [] }));
    assert.strictEqual(toolFailureClaim({ userText: "Soạn hợp đồng", answer, wrote: false, toolsCalled: ["word_create_admin_document"], failedWrites: 1 }), null);
    assert.strictEqual(toolFailureClaim({ userText: "x", answer: "Đã dựng xong văn bản.", wrote: true, toolsCalled: [] }), null);
  });
  it("'soạn nghị quyết', 'biên bản họp đại hội đồng cổ đông' bật nhóm công cụ hành chính và kỹ năng", () => {
    for (const q of ["Soạn nghị quyết của Hội đồng quản trị", "Lập biên bản họp Đại hội đồng cổ đông bất thường"]) {
      assert.ok(requestGroups(WORD_TOOL_GROUPS, { text: q }).has("hanh_chinh"), q);
    }
    const sk = BUILTIN_SKILLS.find((s) => s.name === "soan-thao-van-ban-hanh-chinh")!;
    assert.ok(sk.triggers!.includes("nghi quyet") && sk.body.includes("QUYẾT NGHỊ:") && sk.body.includes("Điều 145"));
  });
  it("kỹ năng soạn có danh mục căn cứ đang có hiệu lực và văn bản đã hết hiệu lực, không ghi số hiệu", () => {
    for (const name of ["soan-thao-van-ban-hanh-chinh", "soan-hop-dong"]) {
      const b = BUILTIN_SKILLS.find((s) => s.name === name)!.body;
      assert.ok(b.includes("Luật Nhà ở năm 2023") && b.includes("ĐÃ HẾT HIỆU LỰC") && b.includes("Luật Đất đai 2013"), name);
      assert.ok(!/\b\d{1,3}\/\d{4}\/[A-ZĐ]/.test(b), name);
    }
  });
});

describe("4.11.10 · Vòng lọc OpenRouter (24/11/2026)", () => {
  it("'\n' viết bằng chữ lẫn giữa dòng thật cũng tách; '.;' cuối dòng căn cứ thành ';'", () => {
    assert.deepStrictEqual(bodyLinesOf("# QUYẾT ĐỊNH:\nDòng A\n## Điều 1. Bổ nhiệm"), ["# QUYẾT ĐỊNH:", "Dòng A", "## Điều 1. Bổ nhiệm"]);
    assert.deepStrictEqual(bodyLinesOf("_Căn cứ Điều lệ công ty.;_"), ["_Căn cứ Điều lệ công ty;_"]);
  });
  it("chữ mẫu của kỹ năng chép nguyên vào văn bản: báo", () => {
    const p = planAdminDocument({ doc_type: "quyet_dinh", org: ORG, body: "# THẨM QUYỀN BAN HÀNH\n## Điều 1. Bổ nhiệm …" });
    assert.ok(p.checks?.some((c) => c.includes("chữ mẫu")), String(p.checks));
  });
  it("tên đơn vị, người ký, cụm chuẩn viết sai dấu: báo; viết đúng thì không", () => {
    const p = planAdminDocument({
      doc_type: "quyet_dinh",
      org: ORG,
      body: "_Căn cứ Điều lệ Công ty Cổ phẫn Đầu tư Anh Quân;_\n## Điều 1. Hội đòng quản trị bổ nhiệm ông Trần Văn Bình.",
      signer_title: "CHỦ TỊCH",
      signer_name: "Nguyễn Anh Quân",
    });
    const c = (p.checks || []).join(" ");
    assert.ok(c.includes("Cổ phẫn Đầu tư Anh Quân") && c.includes("Hội đòng quản trị"), c);
    const ok = planAdminDocument({ doc_type: "quyet_dinh", org: ORG, body: "_Căn cứ Điều lệ Công ty Cổ phần Đầu tư Anh Quân;_\n## Điều 1. Hội đồng quản trị bổ nhiệm ông Bình.", signer_authority: "TM. HỘI ĐỒNG QUẢN TRỊ", signer_title: "CHỦ TỊCH" });
    assert.strictEqual(ok.checks, undefined, String(ok.checks));
  });
});

describe("4.11.10 · Kiểm chứng trên bản cục bộ (24/11/2026)", () => {
  it("một cặp _…_ bao nhiều dòng căn cứ: nghiêng từng dòng, không còn dấu gạch dưới", () => {
    assert.deepStrictEqual(bodyLinesOf("_Căn cứ Điều lệ;\nCăn cứ Biên bản họp;\nCăn cứ kết quả biểu quyết._\n# QUYẾT NGHỊ:"), [
      "_Căn cứ Điều lệ;_",
      "_Căn cứ Biên bản họp;_",
      "_Căn cứ kết quả biểu quyết._",
      "# QUYẾT NGHỊ:",
    ]);
    // gạch đầu dòng bằng gạch dưới không bị coi là mở nghiêng
    assert.deepStrictEqual(bodyLinesOf("_ Số tiền vay: 150 tỷ đồng.\n_Căn cứ Điều lệ;_"), ["- Số tiền vay: 150 tỷ đồng.", "_Căn cứ Điều lệ;_"]);
  });
  it("'Hộp đại hội đồng cổ đông', 'NGHỊ SỨ' bị báo sai dấu", () => {
    const p = planAdminDocument({ doc_type: "bien_ban", org: ORG, subject: "Hộp đại hội đồng cổ đông bất thường năm 2026", body: "## NGHỊ SỨ\n- Thông qua chủ trương đầu tư." });
    const c = (p.checks || []).join(" ");
    assert.ok(c.includes("Hộp đại hội đồng cổ đông") && c.includes("NGHỊ SỨ"), c);
  });
});

describe("4.11.10 · Bổ sung sau kiểm chứng Cohere", () => {
  it("dẫn văn bản đã hết hiệu lực: báo trong checks và dòng tóm tắt", () => {
    const p = planAdminDocument({ doc_type: "cong_van", org: ORG, subject: "Đề nghị", to: ["Sở Xây dựng"], body: "Căn cứ Luật Nhà ở năm 2023, Nghị định số 99/2015/NĐ-CP và các văn bản hướng dẫn, Công ty kính đề nghị." });
    assert.ok(p.checks?.some((c) => c.includes("HẾT HIỆU LỰC") && c.includes("99/2015")), String(p.checks));
    assert.ok(!p.checks?.join(" ").includes("Luật Nhà ở 2014 (thay"));
    assert.ok(adminDocBrief(p).includes("chỗ nghi sai cần kiểm"));
  });
  it("họ tên người ký toàn chữ hoa đổi về chữ thường in đậm", () => {
    const p = planAdminDocument({ doc_type: "quyet_dinh", org: ORG, body: "## Điều 1. …", signer_title: "CHỦ TỊCH", signer_name: "NGUYỄN ANH QUÂN" });
    assert.ok(allText(p).includes("Nguyễn Anh Quân"));
  });
});

describe("4.11.10 · Dòng nghiêng mở mà không đóng", () => {
  it("'_Căn cứ …;' không có dấu đóng: vẫn là dòng nghiêng, không in dấu gạch dưới", () => {
    assert.deepStrictEqual(bodyLinesOf("_Căn cứ Luật Nhà ở năm 2023;\n_Căn cứ [cần bổ sung căn cứ]."), ["_Căn cứ Luật Nhà ở năm 2023;_", "_Căn cứ [cần bổ sung căn cứ]._"]);
    const p = planAdminDocument({ doc_type: "hop_dong", title: "HỢP ĐỒNG CHUYỂN NHƯỢNG", body: "_Căn cứ Luật Nhà ở năm 2023;\nĐiều 1. Đối tượng" });
    const c = paras(p).find((x) => x.text.startsWith("Căn cứ"))!;
    assert.ok(c.italic && !c.text.includes("_"));
  });
});
