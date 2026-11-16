/** Đợt 3 — phần A: kỹ năng và công cụ soạn văn bản hành chính; cảnh báo "nói đã sửa mà không ghi". */
import * as assert from "assert";
import { ADMIN_PAGE, DOC_TYPES, ParaBlock, TableBlock, adminDocBrief, estimateWidthMm, planAdminDocument } from "../../src/core/adminDoc";
import { BUILTIN_SKILLS, autoSelectSkills } from "../../src/core/skillpack";
import { resolveSlash } from "../../src/core/skills";
import { WORD_SPECS } from "../../src/office/wordTools";
import { QUICK_PROMPTS, describeCall } from "../../src/office/tools";

const tables = (p: ReturnType<typeof planAdminDocument>) => p.blocks.filter((b) => b.kind === "table") as TableBlock[];
const paras = (p: ReturnType<typeof planAdminDocument>) => p.blocks.filter((b) => b.kind === "para") as ParaBlock[];
const cells = (t: TableBlock) => t.rows.flatMap((r) => [...r[0], ...r[1]]);

describe("Đợt 3 · A1 · Bản vẽ văn bản hành chính", () => {
  const base = {
    doc_type: "cong_van",
    org: "Công ty cổ phần Đầu tư ABC",
    parent_org: "Tổng công ty XYZ",
    symbol: "ABC-PC",
    place: "Hà Nội",
    subject: "đề nghị gia hạn tiến độ dự án",
    to: ["Sở Xây dựng thành phố Hà Nội"],
    body: "Công ty kính gửi Quý Sở lời chào trân trọng.\nCông ty đề nghị gia hạn 06 tháng.\nTrân trọng.",
    signer_title: "Giám đốc",
    signer_name: "Nguyễn Văn A",
    year: 2026,
  };

  it("Công văn: đủ thành phần, đúng cỡ chữ và kiểu chữ", () => {
    const p = planAdminDocument(base);
    const [head, sign] = tables(p);
    const h = cells(head);
    const by = (t: string) => h.find((c) => c.text.includes(t))!;
    assert.deepStrictEqual({ size: by("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM").size, bold: by("CỘNG HÒA").bold }, { size: 13, bold: true });
    assert.deepStrictEqual({ size: by("Độc lập - Tự do - Hạnh phúc").size, bold: by("Độc lập").bold }, { size: 14, bold: true });
    assert.ok(by("TỔNG CÔNG TY XYZ") && !by("TỔNG CÔNG TY XYZ").bold, "cơ quan chủ quản in hoa, không đậm");
    assert.ok(by("CÔNG TY CỔ PHẦN ĐẦU TƯ ABC").bold, "cơ quan ban hành in hoa, đậm");
    assert.strictEqual(by("Số:").text, "Số:      /ABC-PC");
    assert.strictEqual(by("Hà Nội,").text, "Hà Nội, ngày      tháng      năm 2026");
    assert.ok(by("Hà Nội,").italic);
    assert.strictEqual(by("V/v").text, "V/v đề nghị gia hạn tiến độ dự án");
    assert.strictEqual(h.filter((c) => c.ruleMm).length, 2, "hai đường kẻ: dưới tên cơ quan và dưới tiêu ngữ");
    assert.strictEqual(head.widths[0] + head.widths[1], 210 - ADMIN_PAGE.leftMm - ADMIN_PAGE.rightMm);
    const ps = paras(p);
    assert.ok(ps.some((x) => x.text === "Kính gửi: Sở Xây dựng thành phố Hà Nội"));
    const body = ps.filter((x) => x.indent);
    assert.strictEqual(body.length, 3);
    assert.ok(body.every((x) => x.size === 14 && x.align === "Justified" && (x.after || 0) >= 6));
    assert.strictEqual(body[2].text, "Trân trọng./.");
    const s = cells(sign);
    assert.ok(s.find((c) => c.text === "GIÁM ĐỐC")!.bold);
    assert.ok(s.find((c) => c.text === "Nguyễn Văn A")!.bold);
    const nn = s.find((c) => c.text === "Nơi nhận:")!;
    assert.deepStrictEqual({ size: nn.size, bold: nn.bold, italic: nn.italic }, { size: 12, bold: true, italic: true });
    assert.deepStrictEqual(
      s.filter((c) => c.text.startsWith("- ")).map((c) => [c.text, c.size]),
      [
        ["- Như trên;", 11],
        ["- Lưu: VT, ….", 11],
      ]
    );
    // Chừa ít nhất bốn dòng cho chữ ký
    const right = sign.rows[0][1].map((c) => c.text);
    assert.ok(right.indexOf("Nguyễn Văn A") - right.indexOf("GIÁM ĐỐC") >= 5);
    assert.deepStrictEqual(p.blanks, ["số văn bản (để văn thư cấp)", "ngày tháng ban hành", "đơn vị soạn thảo ở dòng Lưu"]);
    assert.match(adminDocBrief(p), /công văn.*3 chỗ cần điền/);
  });

  it("Tờ trình: có tên loại, trích yếu đậm, đường kẻ dưới trích yếu; không có V/v", () => {
    const p = planAdminDocument({ ...base, doc_type: "to_trinh", symbol: "TTr-ABC", subject: "Về việc phê duyệt phương án", date: "5/1/2026", font_size: 13 });
    const ps = paras(p);
    const i = ps.findIndex((x) => x.text === "TỜ TRÌNH");
    assert.ok(i >= 0 && ps[i].bold && ps[i].align === "Centered" && ps[i].size === 13);
    assert.ok(ps[i + 1].text === "Về việc phê duyệt phương án" && ps[i + 1].bold);
    assert.ok((ps[i + 2].ruleMm || 0) >= 15);
    const h = cells(tables(p)[0]);
    assert.ok(!h.some((c) => c.text.startsWith("V/v")));
    assert.ok(h.some((c) => c.text === "Hà Nội, ngày 05 tháng 01 năm 2026"));
    // Cỡ 13: cặp quốc hiệu / tiêu ngữ là 12 / 13
    assert.strictEqual(h.find((c) => c.text.includes("CỘNG HÒA"))!.size, 12);
    assert.strictEqual(h.find((c) => c.text.includes("Độc lập"))!.size, 13);
    assert.ok(!p.blanks.includes("ngày tháng ban hành"));
  });

  it("Quyết định: căn cứ nghiêng, đề mục canh giữa, Điều đậm; không tự thêm Như trên", () => {
    const p = planAdminDocument({
      doc_type: "quyet_dinh",
      org: "Công ty ABC",
      symbol: "QĐ-ABC",
      subject: "Về việc thành lập Tổ công tác",
      body: "# GIÁM ĐỐC CÔNG TY ABC\n_Căn cứ Điều lệ Công ty;_\n# QUYẾT ĐỊNH:\n## Điều 1. Thành lập Tổ công tác\nTổ công tác gồm 05 thành viên.\n## Điều 2. Hiệu lực thi hành\nQuyết định có hiệu lực từ ngày ký.",
      signer_title: "GIÁM ĐỐC",
      recipients: ["Như Điều 2", "Lưu: VT, HCNS"],
    });
    const ps = paras(p);
    assert.ok(ps.find((x) => x.text === "Căn cứ Điều lệ Công ty;")!.italic);
    assert.strictEqual(ps.find((x) => x.text === "QUYẾT ĐỊNH:")!.align, "Centered");
    assert.ok(ps.find((x) => x.text.startsWith("Điều 1."))!.bold);
    const s = cells(tables(p)[1]).filter((c) => c.text.startsWith("- ")).map((c) => c.text);
    assert.deepStrictEqual(s, ["- Như Điều 2;", "- Lưu: VT, HCNS."]);
    assert.ok(p.blanks.includes("họ tên người ký") && p.blanks.includes("địa danh"));
    // Không có cơ quan chủ quản: quốc hiệu ngang hàng tên cơ quan, vẫn đủ hai đường kẻ
    assert.strictEqual(cells(tables(p)[0]).filter((c) => c.ruleMm).length, 2);
  });

  it("Biên bản: hai người ký, không có địa danh ngày tháng ở đầu, Nơi nhận nằm dưới", () => {
    const p = planAdminDocument({ doc_type: "bien_ban", org: "Công ty ABC", subject: "Họp giao ban tháng 11", body: "Thời gian bắt đầu: 8 giờ.\nNội dung: …", left_signer_title: "Thư ký", left_signer_name: "Trần Thị B", signer_title: "Chủ tọa", signer_name: "Nguyễn Văn A" });
    const sign = tables(p)[1];
    assert.deepStrictEqual(sign.rows[0].map((c) => c[0].text), ["THƯ KÝ", "CHỦ TỌA"]);
    assert.deepStrictEqual(sign.rows[0].map((c) => c[c.length - 1].text), ["Trần Thị B", "Nguyễn Văn A"]);
    assert.strictEqual(sign.rows[0][0].length, sign.rows[0][1].length);
    assert.ok(!cells(tables(p)[0]).some((c) => /ngày\s+tháng/.test(c.text)));
    assert.ok(paras(p).some((x) => x.text === "- Lưu: VT, Hồ sơ."));
  });

  it("Nhiều nơi nhận ở Kính gửi: mỗi nơi một dòng, chấm phẩy rồi chấm", () => {
    const p = planAdminDocument({ ...base, to: ["Sở Xây dựng", "Sở Tài chính;", "- Ủy ban nhân dân quận"] });
    const ps = paras(p);
    const k = ps.findIndex((x) => x.text === "Kính gửi:");
    assert.deepStrictEqual(ps.slice(k + 1, k + 4).map((x) => x.text), ["- Sở Xây dựng;", "- Sở Tài chính;", "- Ủy ban nhân dân quận."]);
    assert.ok((ps[k + 1].leftMm || 0) > (ps[k].leftMm || 0));
    // Chỉ hai bảng (phần đầu, phần ký) và không bảng nào đứng ngay sau một đoạn rỗng
    assert.strictEqual(tables(p).length, 2);
    p.blocks.forEach((b, i) => assert.ok(!(b.kind === "table" && i > 0 && p.blocks[i - 1].kind === "para" && !(p.blocks[i - 1] as ParaBlock).text), "bảng sau đoạn rỗng"));
    // Đường kẻ không bao giờ là đoạn cuối của một ô (viền của đoạn cuối ô dính sang ô cùng hàng)
    for (const t of tables(p)) for (const r of t.rows) for (const c of r) assert.ok(c.length >= 1 && !c[c.length - 1].ruleMm);
  });

  it("Không bịa: thiếu cơ quan ban hành hoặc nội dung thì từ chối; bỏ học hàm học vị khỏi tên người ký", () => {
    assert.throws(() => planAdminDocument({ doc_type: "cong_van", body: "x" }), /cơ quan/);
    assert.throws(() => planAdminDocument({ doc_type: "cong_van", org: "ABC" }), /nội dung/);
    assert.throws(() => planAdminDocument({ doc_type: "khac", org: "ABC", body: "x" }), /tên loại/);
    const p = planAdminDocument({ ...base, signer_name: "TS. Nguyễn Văn A", signer_authority: "KT. Giám đốc", signer_title: "Phó Giám đốc" });
    const s = cells(tables(p)[1]).map((c) => c.text);
    assert.ok(s.includes("Nguyễn Văn A") && s.includes("KT. GIÁM ĐỐC") && s.includes("PHÓ GIÁM ĐỐC"));
  });

  it("Loại không nhận ra thì coi là 'khac' và cần title; ước lượng bề rộng chữ hợp lý", () => {
    assert.ok(DOC_TYPES.includes("ke_hoach"));
    const p = planAdminDocument({ doc_type: "phuong-an", title: "Phương án", org: "ABC", body: "x" });
    assert.ok(paras(p).some((x) => x.text === "PHƯƠNG ÁN"));
    const w = estimateWidthMm("Độc lập - Tự do - Hạnh phúc", 14, { bold: true });
    assert.ok(w > 50 && w < 66, String(w));
  });
});

describe("Đợt 3 · A1 · Kỹ năng soạn thảo văn bản hành chính", () => {
  const word = BUILTIN_SKILLS.filter((s) => !s.hosts || s.hosts.includes("Word"));
  it("Tự nạp khi yêu cầu nhắc công văn, tờ trình, quyết định, thông báo, giấy mời, biên bản", () => {
    for (const q of [
      "Soạn một công văn gửi Sở Xây dựng đề nghị gia hạn tiến độ",
      "soạn giúp tôi tờ trình xin phê duyệt phương án bồi thường",
      "Dự thảo quyết định thành lập tổ công tác",
      "Soạn thông báo nghỉ lễ cho toàn công ty",
      "Làm giấy mời họp cổ đông",
      "Soạn biên bản họp giao ban",
      "Văn bản này đúng thể thức chưa?",
    ])
      assert.ok(autoSelectSkills(q, word).some((s) => s.name === "soan-thao-van-ban-hanh-chinh"), q);
  });
  it("Không chen vào yêu cầu khác; không cướp chỗ kỹ năng trả lời công văn", () => {
    assert.ok(!autoSelectSkills("Rà soát hợp đồng mua bán này giúp tôi", word).some((s) => s.name === "soan-thao-van-ban-hanh-chinh"));
    assert.deepStrictEqual(autoSelectSkills("xin chào", word), []);
    const both = autoSelectSkills("Soạn công văn trả lời Sở Tài chính", word).map((s) => s.name);
    assert.ok(both.includes("soan-thao-van-ban-hanh-chinh") && both.includes("soan-phan-hoi-phap-ly"), both.join(","));
  });
  it("Chỉ có ở Word; gọi được bằng /", () => {
    const sk = BUILTIN_SKILLS.find((s) => s.name === "soan-thao-van-ban-hanh-chinh")!;
    assert.deepStrictEqual(sk.hosts, ["Word"]);
    const r = resolveSlash("/soan-thao-van-ban-hanh-chinh soạn tờ trình xin chủ trương", BUILTIN_SKILLS, []);
    assert.strictEqual(r.label, "/soan-thao-van-ban-hanh-chinh");
    assert.ok(r.content.includes("word_create_admin_document") && r.content.includes("soạn tờ trình xin chủ trương"));
  });
  it("Kỹ năng ngắn gọn và nhắc đúng công cụ dựng, kiểm, sửa thể thức", () => {
    const sk = BUILTIN_SKILLS.find((s) => s.name === "soan-thao-van-ban-hanh-chinh")!;
    assert.ok(sk.body.length < 9000, `kỹ năng dài ${sk.body.length} ký tự`);
    for (const t of ["word_create_admin_document", "word_check_page_format", "word_set_page_format", "word_format_text"]) assert.ok(sk.body.includes(t), t);
    assert.ok(!/\b\d{1,3}\/\d{4}\/[A-ZĐ]/.test(sk.body + sk.description), "không ghi số hiệu văn bản pháp luật");
  });
  it("Hai công cụ mới có trong bộ công cụ Word, là công cụ ghi, có nhãn tiếng Việt; màn chào có gợi ý soạn công văn", () => {
    for (const n of ["word_create_admin_document", "word_set_page_format"]) {
      const sp = WORD_SPECS.find((s) => s.name === n)!;
      assert.ok(sp && sp.write, n);
      assert.notStrictEqual(describeCall({ id: "x", name: n, args: {} }), n);
    }
    assert.ok(QUICK_PROMPTS.Word.some((q) => /công văn/.test(q)));
    const fmt = WORD_SPECS.find((s) => s.name === "word_format_text")!;
    assert.ok(fmt.parameters.properties.first_line_indent_mm && fmt.parameters.properties.line_spacing);
  });
});

import { REDO_WITH_TOOLS, findWriteClaim, hasEditIntent, unverifiedWriteClaim } from "../../src/core/claims";

describe("Đợt 3 · A2 · Cảnh báo 'nói đã sửa mà không ghi'", () => {
  const ask = "Sửa mức phạt ở Điều 5 từ 150% xuống 120%";
  const warn = (answer: string, o: Partial<Parameters<typeof unverifiedWriteClaim>[0]> = {}) =>
    unverifiedWriteClaim({ userText: ask, answer, wrote: false, toolsCalled: [], ...o });

  it("Khẳng định đã làm mà không có thao tác ghi: cảnh báo", () => {
    for (const a of [
      "Đã sửa mức phạt tại Điều 5 từ 150% xuống 120%.",
      "Tôi đã thay “150%” bằng “120%” ở Điều 5.2.",
      "✅ Đã cập nhật Điều 5.",
      "Xong rồi ạ. Em đã chỉnh sửa điều khoản theo yêu cầu.",
      "Điều 5.2 đã được sửa thành 120%.",
      "- Đã thay 150% bằng 120% ở [¶57]\n- Đã thêm nhận xét giải thích",
      "**Đã sửa xong**, anh có thể kiểm tra lại ở Điều 5.",
      "I've updated clause 5 to 120%.",
      "The penalty rate has been changed to 120%.",
      "Văn bản đã được cập nhật theo yêu cầu của anh.",
      "Tôi vừa chèn công văn vào văn bản.",
    ])
      assert.ok(warn(a), a);
  });

  it("Có ghi thật thì không cảnh báo, dù câu trả lời nói đã sửa", () => {
    assert.strictEqual(warn("Đã sửa mức phạt tại Điều 5.", { wrote: true, toolsCalled: ["word_replace_text"] }), null);
  });

  it("Mô tả, đề xuất, hỏi lại, nói chưa làm: không cảnh báo", () => {
    for (const a of [
      "Điều 5 quy định mức phạt 150%. Anh có muốn tôi sửa thành 120% không?",
      "Tôi sẽ sửa Điều 5 từ 150% xuống 120% nếu anh đồng ý.",
      "Tôi chưa sửa được vì không tìm thấy Điều 5 trong văn bản.",
      "Tôi không thể sửa trực tiếp; đề xuất câu chữ như sau: “Mức phạt là 120%”.",
      "Đề xuất: thay 150% bằng 120% ở Điều 5.2.",
      "Nếu đã sửa mức phạt thì cần sửa cả Điều 9 cho thống nhất.",
      "Anh nên sửa Điều 5 như sau: mức phạt 120%.",
      "Sau khi anh xác nhận, tôi sẽ thay ngay.",
      "Mức phạt hiện là 150%; để sửa, tôi cần biết anh muốn áp dụng từ ngày nào.",
      "Văn bản có câu “Bên B đã thay đổi địa chỉ” ở Điều 3.",
      "Đã đọc Điều 5: mức phạt đang là 150%.",
      "Đã kiểm tra văn bản, không thấy Điều 5.",
      "Bạn đã sửa mức phạt ở lượt trước rồi, hiện là 120%.",
      "```\nĐã sửa Điều 5\n```\nĐây là câu mẫu để ghi vào biên bản.",
    ])
      assert.strictEqual(warn(a), null, a);
  });

  it("Kể lại sửa đổi của đối tác (đã gọi công cụ đọc sửa đổi), người dùng không nhờ sửa: không cảnh báo", () => {
    const f = { userText: "Đối tác đã đụng vào những điều nào?", wrote: false, toolsCalled: ["word_list_revisions"] };
    assert.strictEqual(unverifiedWriteClaim({ ...f, answer: "Điều 5.2 đã được sửa từ 30 thành 45 ngày.\n- Đã thêm khoản 8.3 về bảo hành." }), null);
    assert.strictEqual(unverifiedWriteClaim({ ...f, answer: "Đã xóa Điều 12 và đã đổi tên Bên B ở phần đầu." }), null);
    // Nhưng người dùng nhờ chấp nhận mà AI nói đã chấp nhận, không gọi công cụ duyệt: cảnh báo
    assert.ok(unverifiedWriteClaim({ userText: "Chấp nhận các sửa đổi ở Điều 5", answer: "Đã chấp nhận 3 sửa đổi ở Điều 5.", wrote: false, toolsCalled: ["word_list_revisions"] }));
  });

  it("Người dùng chỉ hỏi, AI dùng câu bị động để mô tả: không cảnh báo; AI tự nhận 'tôi đã sửa': vẫn cảnh báo", () => {
    assert.strictEqual(unverifiedWriteClaim({ userText: "Điều 5 nói gì?", answer: "Điều 5 đã được sửa đổi theo phụ lục 02.", wrote: false, toolsCalled: ["word_read_clause"] }), null);
    assert.ok(unverifiedWriteClaim({ userText: "Điều 5 nói gì?", answer: "Tôi đã sửa Điều 5 cho rõ hơn.", wrote: false, toolsCalled: ["word_read_clause"] }));
  });

  it("Câu cảnh báo nêu đúng lý do: không gọi công cụ / chỉ đọc / ghi đều lỗi", () => {
    assert.match(warn("Đã sửa Điều 5.")!.text, /không gọi công cụ nào.*CHƯA được thay đổi/);
    assert.match(warn("Đã sửa Điều 5.", { toolsCalled: ["word_find"] })!.text, /chỉ gọi công cụ đọc/);
    assert.match(warn("Đã sửa Điều 5.", { toolsCalled: ["word_replace_text"], failedWrites: 2 })!.text, /2 thao tác ghi của lượt này đều không thành công/);
    assert.ok(warn("Đã sửa Điều 5.")!.text.includes("Đã sửa Điều 5"));
  });

  it("Nhận diện từng phần: độ mạnh của lời khẳng định, ý nhờ sửa của người dùng", () => {
    assert.strictEqual(findWriteClaim("Tôi đã thêm khoản 8.3.").strength, "strong");
    assert.strictEqual(findWriteClaim("Đã định dạng đoạn 3 thành chữ đậm.").strength, "medium");
    assert.strictEqual(findWriteClaim("Điều 5 nói về phạt vi phạm.").strength, "none");
    assert.strictEqual(findWriteClaim("Điều 5.3 đã được sửa.").strength, "medium");
    assert.ok(hasEditIntent("soạn giúp tôi một công văn"));
    assert.ok(hasEditIntent("Thêm một khoản sau 8.2"));
    assert.ok(!hasEditIntent("Điều 5 nói gì?"));
    assert.ok(REDO_WITH_TOOLS.includes("GỌI CÔNG CỤ"));
  });
});

import { isToolUnsupportedError } from "../../src/core/agent/llm";

describe("Đợt 3 · A3 · Model không gọi được công cụ", () => {
  it("Câu lỗi của Groq (16/11/2026) được nhận là model không gọi được công cụ", () => {
    assert.ok(isToolUnsupportedError({ status: 400, message: "Lỗi 400 (`tool calling` is not supported with this model)" }));
    assert.ok(isToolUnsupportedError({ status: 400, message: "This model does not support tools" }));
    assert.ok(!isToolUnsupportedError({ status: 400, message: "Invalid request: messages is empty" }));
  });
});
