/** Bản 4.11.8: soạn hợp đồng đúng thể thức (Chạy thật 23/11/2026 với Nemotron: hợp đồng dựng bằng loại 'khac' thừa dòng tên
 *  hợp đồng, có Nơi nhận và "./."; sửa phần đầu thì model xóa tên đơn vị ở ô trái rồi chèn xuống dưới tiêu ngữ ở ô phải). */
import * as assert from "assert";
import { DOC_TYPES, ParaBlock, TableBlock, adminDocBrief, planAdminDocument, stripHeaderEcho } from "../../src/core/adminDoc";
import { BUILTIN_SKILLS, autoSelectSkills } from "../../src/core/skillpack";
import { resolveSlash } from "../../src/core/skills";
import { WORD_TOOL_GROUPS, requestGroups } from "../../src/core/toolgroups";
import { TABLE_NOTE, packParagraphs } from "../../src/office/util";
import { WORD_SPECS } from "../../src/office/wordTools";

const tables = (p: ReturnType<typeof planAdminDocument>) => p.blocks.filter((b) => b.kind === "table") as TableBlock[];
const paras = (p: ReturnType<typeof planAdminDocument>) => p.blocks.filter((b) => b.kind === "para") as ParaBlock[];
const allText = (p: ReturnType<typeof planAdminDocument>) =>
  p.blocks.flatMap((b) => (b.kind === "table" ? b.rows.flatMap(([l, r]) => [...l, ...r]).map((x) => x.text) : [b.text]));

/** Nội dung như model đã gửi ngày 23/11/2026: chép lại phần đầu ở đầu body và phần ký ở cuối body */
const ECHO_BODY = [
  "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
  "Độc lập - Tự do - Hạnh phúc",
  "---------------",
  "**HỢP ĐỒNG XÂY DỰNG**",
  "Số: …/2026/HĐ-XD",
  "Hà Nội, ngày … tháng … năm 2026",
  "_Căn cứ Bộ luật Dân sự;_",
  "_Căn cứ nhu cầu và khả năng của hai bên._",
  "Hôm nay, ngày … tháng … năm 2026, chúng tôi gồm:",
  "## BÊN GIAO THẦU (BÊN A): CÔNG TY CỔ PHẦN ABC",
  "Đại diện: … Chức vụ: …",
  "## BÊN NHẬN THẦU (BÊN B): …",
  "Hai bên thống nhất ký kết hợp đồng với các điều khoản sau:",
  "## Điều 1. Nội dung công việc",
  "Bên B thi công xây dựng công trình theo thiết kế được duyệt.",
  "## Điều 2. Điều khoản chung",
  "Hợp đồng lập thành 04 bản có giá trị pháp lý như nhau, mỗi bên giữ 02 bản.",
  "ĐẠI DIỆN BÊN A",
  "ĐẠI DIỆN BÊN B",
].join("\n");

describe("4.11.8 · Hợp đồng: bản vẽ loại hop_dong", () => {
  it("Có loại hop_dong; tên loại nhận cả 'hợp đồng', 'hop_dong_xay_dung'", () => {
    assert.ok((DOC_TYPES as readonly string[]).includes("hop_dong"));
    for (const t of ["hop_dong", "Hợp đồng", "hợp đồng xây dựng", "hop-dong-thi-cong"])
      assert.strictEqual(planAdminDocument({ doc_type: t, title: "Hợp đồng thi công xây dựng", body: "## Điều 1. Nội dung" }).type, "hop_dong", t);
  });

  it("Có tên đơn vị: phần đầu hai cột — trái đơn vị và Số, phải quốc hiệu, tiêu ngữ; tên hợp đồng chỉ một dòng", () => {
    const p = planAdminDocument({ doc_type: "hop_dong", org: "Công ty cổ phần ABC", title: "Hợp đồng thi công xây dựng", symbol: "HĐ-XD", body: ECHO_BODY });
    const [head, sign] = tables(p);
    assert.ok(head && sign, "phải có bảng phần đầu và bảng phần ký");
    const [left, right] = head.rows[0];
    assert.ok(left.some((x) => x.text === "CÔNG TY CỔ PHẦN ABC" && x.bold), "tên đơn vị in hoa đậm ở ô trái");
    assert.strictEqual(right[0].text, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    assert.strictEqual(right[1].text, "Độc lập - Tự do - Hạnh phúc");
    assert.strictEqual(head.rows.length, 1, "không có ngày ở đầu: một hàng (hàng có ô rỗng làm Hoàn tác sót đoạn)");
    assert.ok(left[left.length - 1].text.startsWith("Số: "), "Số nằm cuối ô trái, dưới tên đơn vị");
    const dated = tables(planAdminDocument({ doc_type: "hop_dong", org: "ABC", title: "Hợp đồng", place: "Hà Nội", body: "x" }))[0];
    assert.ok(dated.rows.length === 2 && dated.rows[1][0][0].text.startsWith("Số: ") && /^Hà Nội, ngày/.test(dated.rows[1][1][0].text));
    assert.ok(head.rows[0][0].some((x) => x.ruleMm) && head.rows[0][1].some((x) => x.ruleMm), "đường kẻ dưới đơn vị và tiêu ngữ");
    const texts = allText(p);
    assert.strictEqual(texts.filter((t) => /^HỢP ĐỒNG/.test(t)).length, 1, texts.join(" | "));
    assert.strictEqual(texts.filter((t) => t === "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM").length, 1);
    assert.ok(!texts.some((t) => /^Số: …\/2026/.test(t)), "dòng Số chép trong body đã bỏ");
    assert.ok(p.notes.some((n) => /đã bỏ \d+ dòng/.test(n)), p.notes.join("; "));
  });

  it("Không Kính gửi, không Nơi nhận, không './.'; căn cứ nghiêng; câu 'Hôm nay, ngày …' giữ nguyên", () => {
    const p = planAdminDocument({ doc_type: "hop_dong", title: "HỢP ĐỒNG THI CÔNG", body: ECHO_BODY, to: ["Sở Xây dựng"], recipients: ["Lưu: VT"] });
    const texts = allText(p);
    assert.ok(!texts.some((t) => /Kính gửi|Nơi nhận|Lưu: VT/.test(t)), texts.join(" | "));
    assert.ok(!texts.some((t) => /\.\/\.$/.test(t)), "không có dấu ./.");
    assert.ok(paras(p).some((x) => x.text === "Căn cứ Bộ luật Dân sự;" && x.italic));
    const q = planAdminDocument({ doc_type: "hop_dong", title: "Hợp đồng", body: ["Căn cứ nhu cầu hai bên;", "## Điều 1. Căn cứ tính giá", "Căn cứ vào khối lượng thực tế."].join("\n") });
    const qs = paras(q);
    assert.ok(qs.find((x) => x.text === "Căn cứ nhu cầu hai bên;")!.italic, "căn cứ không đánh dấu vẫn nghiêng");
    assert.ok(!qs.find((x) => x.text === "Điều 1. Căn cứ tính giá")!.italic, "đề mục đậm không nghiêng");
    assert.ok(texts.includes("Hôm nay, ngày … tháng … năm 2026, chúng tôi gồm:"));
    assert.ok(p.notes.some((n) => /Kính gửi và Nơi nhận/.test(n)));
  });

  it("Không tên đơn vị: quốc hiệu, tiêu ngữ canh giữa, chỉ một bảng (phần ký); Số dưới tên hợp đồng", () => {
    const p = planAdminDocument({ doc_type: "hop_dong", title: "Hợp đồng tư vấn giám sát", number: "01/2026/HĐTV", body: "## Điều 1. Nội dung" });
    assert.strictEqual(tables(p).length, 1);
    const ps = paras(p);
    assert.strictEqual(ps[0].text, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    assert.strictEqual(ps[0].align, "Centered");
    const ti = ps.findIndex((x) => x.text === "HỢP ĐỒNG TƯ VẤN GIÁM SÁT");
    assert.ok(ti > 0 && ps[ti].bold);
    assert.strictEqual(ps[ti + 1].text, "Số: 01/2026/HĐTV", "số đủ dạng giữ nguyên, không cắt ở dấu /");
    assert.ok(!p.blanks.includes("số hợp đồng"));
  });

  it("Phần ký hai bên: ĐẠI DIỆN BÊN A trái, BÊN B phải, họ tên cùng hàng; thiếu tên thì báo chỗ trống", () => {
    const p = planAdminDocument({ doc_type: "hop_dong", title: "Hợp đồng", body: "x", left_signer_title: "Tổng giám đốc", signer_title: "Giám đốc", signer_name: "Trần Văn B" });
    const sign = tables(p).pop()!;
    const [l, r] = sign.rows[0];
    assert.strictEqual(l[0].text, "ĐẠI DIỆN BÊN A");
    assert.strictEqual(r[0].text, "ĐẠI DIỆN BÊN B");
    assert.strictEqual(l[1].text, "TỔNG GIÁM ĐỐC");
    assert.strictEqual(l.length, r.length, "hai cột cao bằng nhau");
    assert.strictEqual(r[r.length - 1].text, "Trần Văn B");
    assert.ok(p.blanks.includes("họ tên người ký Bên A") && !p.blanks.includes("họ tên người ký Bên B"));
    assert.ok(/hợp đồng/.test(adminDocBrief(p)));
    const q = planAdminDocument({ doc_type: "hop_dong", title: "Hợp đồng", body: "x", left_signer_authority: "Đại diện chủ đầu tư" });
    assert.strictEqual(tables(q).pop()!.rows[0][0][0].text, "ĐẠI DIỆN CHỦ ĐẦU TƯ");
  });

  it("Không cần org; thiếu nội dung thì từ chối; tên không có chữ 'HỢP ĐỒNG' thì tự thêm", () => {
    assert.throws(() => planAdminDocument({ doc_type: "hop_dong", title: "x", body: ECHO_BODY.split("\n").slice(0, 6).join("\n") }), /nội dung hợp đồng/);
    const p = planAdminDocument({ doc_type: "hop_dong", title: "Thi công xây dựng", body: "x" });
    assert.ok(allText(p).includes("HỢP ĐỒNG THI CÔNG XÂY DỰNG"));
    assert.ok(planAdminDocument({ doc_type: "hop_dong", body: "x" }).blanks.includes("tên hợp đồng"));
  });

  it("Văn bản hành chính: dòng lặp phần đầu ở đầu body cũng bị bỏ; dòng giữa body không bao giờ bị bỏ", () => {
    const p = planAdminDocument({ doc_type: "to_trinh", org: "Công ty ABC", subject: "Về việc phê duyệt", body: "TỜ TRÌNH\nVề việc phê duyệt\n## I. Sự cần thiết\nTỜ TRÌNH\nKính trình xem xét." });
    const ps = paras(p).map((x) => x.text);
    assert.strictEqual(ps.filter((t) => t === "TỜ TRÌNH").length, 2, "một dòng tên loại của công cụ + một dòng ở giữa body (giữ nguyên)");
    assert.strictEqual(stripHeaderEcho(["Căn cứ Luật Xây dựng;", "Số: 1/HĐ"], {}).removed, 0, "dòng Số ở cuối không phải phần ký");
    assert.strictEqual(stripHeaderEcho(["Hà Nội, ngày 05 tháng 01 năm 2026", "Nội dung"], {}).removed, 1);
    assert.strictEqual(stripHeaderEcho(["Hôm nay, ngày 05 tháng 01 năm 2026", "Nội dung"], {}).removed, 0);
  });
});

describe("4.11.8 · Chữ lạ lẫn trong nội dung (Nemotron chèn chữ Hán, Ả Rập — Chạy thật 23/11/2026)", () => {
  it("Báo cụm chữ Hán, Ả Rập kèm chữ đứng trước; tiếng Việt có dấu, ký hiệu … § ¶ thì không báo", () => {
    const p = planAdminDocument({ doc_type: "hop_dong", title: "Hợp đồng", body: ["## Điều 15. Giải quyết tranh chấp", "Hai bên có quyền 提交仲裁 tại trọng tài.", "Bên B phải رفع到 hồ sơ."].join("\n") });
    assert.deepStrictEqual(p.foreign, ["Hai bên có quyền 提交仲裁", "Bên B phải رفع到"]);
    const ok = planAdminDocument({ doc_type: "cong_van", org: "Công ty ABC", body: "Đề nghị Quý Sở xem xét … § 2, mục ¶ 3; giá 1.250.000.000 đồng (Một tỷ hai trăm năm mươi triệu đồng)." });
    assert.strictEqual(ok.foreign, undefined);
  });
});

describe("4.11.8 · Đoạn trong ô bảng được đánh dấu khi đọc", () => {
  it("cell đi kèm đoạn trong ô và có lời dặn hai ô cùng hàng nằm ngang nhau", () => {
    const all = [
      { text: "CÔNG TY ABC", style: "Normal", cell: "bảng 1, hàng 1, cột 1" },
      { text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", style: "Normal", cell: "bảng 1, hàng 1, cột 2" },
      { text: "HỢP ĐỒNG", style: "Normal" },
    ];
    const r: any = packParagraphs(all, 0, 3);
    assert.strictEqual(r.paragraphs[0].cell, "bảng 1, hàng 1, cột 1");
    assert.strictEqual(r.paragraphs[2].cell, undefined);
    assert.strictEqual(r.table_note, TABLE_NOTE);
    assert.ok(/NGANG/.test(TABLE_NOTE) && /không xóa đoạn rồi chèn lại/.test(TABLE_NOTE));
    assert.strictEqual((packParagraphs([{ text: "a", style: "Normal" }], 0, 1) as any).table_note, undefined);
  });
});

describe("4.11.8 · Kỹ năng soan-hop-dong", () => {
  const word = BUILTIN_SKILLS.filter((s) => !s.hosts || s.hosts.includes("Word"));
  const picked = (q: string) => autoSelectSkills(q, word).map((s) => s.name);
  it("Tự nạp khi soạn hợp đồng; không chen vào rà soát, điền hợp đồng", () => {
    for (const q of ["Soạn mẫu hợp đồng xây dựng theo chuẩn văn bản hành chính", "Dự thảo hợp đồng thi công gói thầu số 3", "soạn thảo hợp đồng tư vấn giám sát", "Lập hợp đồng mua bán vật tư"])
      assert.ok(picked(q).includes("soan-hop-dong"), q);
    for (const q of ["Rà soát hợp đồng mua bán này giúp tôi", "Điền hợp đồng theo thông tin trong tệp", "Tóm tắt Điều 5 của hợp đồng"])
      assert.ok(!picked(q).includes("soan-hop-dong"), q);
  });
  it("Kỹ năng chỉ đúng công cụ và loại hop_dong; nhóm công cụ hành chính kéo theo ở chế độ Miễn phí", () => {
    const sk = BUILTIN_SKILLS.find((s) => s.name === "soan-hop-dong")!;
    assert.deepStrictEqual(sk.hosts, ["Word"]);
    assert.ok(sk.body.includes('doc_type "hop_dong"') && /KHÔNG chép vào body/.test(sk.body) && /NGANG/.test(sk.body));
    assert.ok(!/\b\d{1,3}\/\d{4}\/[A-ZĐ]/.test(sk.body + sk.description), "không ghi số hiệu văn bản pháp luật");
    const grp = WORD_TOOL_GROUPS.find((g) => g.id === "hanh_chinh")!;
    assert.ok(grp.skills?.includes("soan-hop-dong"));
    assert.ok(requestGroups(WORD_TOOL_GROUPS, { text: "làm tiếp", skills: ["soan-hop-dong"] }).has("hanh_chinh"));
    assert.ok(requestGroups(WORD_TOOL_GROUPS, { text: "Soạn mẫu hợp đồng xây dựng" }).has("hanh_chinh"));
    const r = resolveSlash("/soan-hop-dong hợp đồng thi công nhà xưởng", BUILTIN_SKILLS, []);
    assert.ok(r.content.includes("word_create_admin_document"));
  });
  it("Kỹ năng hành chính chỉ hợp đồng sang loại hop_dong và dặn không sắp lại phần đầu", () => {
    const sk = BUILTIN_SKILLS.find((s) => s.name === "soan-thao-van-ban-hanh-chinh")!;
    assert.ok(/doc_type hop_dong/.test(sk.body) && /soan-hop-dong/.test(sk.body) && /NGANG/.test(sk.body));
  });
  it("Công cụ dựng văn bản nhận hop_dong, không bắt buộc org", () => {
    const sp = WORD_SPECS.find((s) => s.name === "word_create_admin_document")!;
    assert.ok((sp.parameters.properties.doc_type.enum as string[]).includes("hop_dong"));
    assert.ok(!(sp.parameters.required as string[]).includes("org"));
    assert.ok(sp.parameters.properties.left_signer_authority);
  });
});
