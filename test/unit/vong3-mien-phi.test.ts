/** Bản 4.11.11: sửa theo vòng thử thứ ba với model miễn phí (24/11/2026, đề mới: giấy mời, báo cáo, hợp đồng đặt cọc,
 *  hợp đồng lao động, giấy ủy quyền, sửa văn bản có sẵn — `tasks/thu-mien-phi-vong3.md`). Mỗi ca lấy đúng dạng tham số model đã gửi. */
import * as assert from "assert";
import { ParaBlock, TableBlock, partyIsPerson, planAdminDocument } from "../../src/core/adminDoc";
import { legalThresholds } from "../../src/core/legalcheck";
import { issueAfterEvent } from "../../src/core/textcheck";

const ORG = "Công ty Cổ phần Đầu tư Anh Quân";

describe("4.11.11 · Vòng thử 3 (model miễn phí, đề mới)", () => {
  it("giấy mời ghi ngày ban hành trùng ngày họp (dots-3-note): báo trong check_now", () => {
    const p = planAdminDocument({
      doc_type: "giay_moi",
      org: ORG,
      date: "10/12/2026",
      body: "Trân trọng kính mời Quý đại diện tới dự cuộc họp.\nThời gian: 14h00, Thứ Năm, ngày 10/12/2026.\nĐịa điểm: Phòng họp tầng 3.",
      signer_authority: "KT. TỔNG GIÁM ĐỐC",
      signer_title: "PHÓ TỔNG GIÁM ĐỐC",
      signer_name: "Trần Văn Bình",
    });
    assert.ok(p.checks?.some((c) => /trùng ngày họp 10\/12\/2026/.test(c)), String(p.checks));
  });
  it("giấy mời ban hành trước ngày họp, hoặc chưa có ngày ban hành: không báo", () => {
    assert.strictEqual(issueAfterEvent("ngày 01 tháng 12 năm 2026", ["Thời gian: 14h00, Thứ Năm, ngày 10/12/2026."]), null);
    assert.strictEqual(issueAfterEvent("ngày      tháng      năm 2026", ["Thời gian: 14h00, Thứ Năm, ngày 10/12/2026."]), null);
    // Ngày trong dòng không phải giờ họp (hạn đăng ký, ngày căn cứ) không tính là ngày họp
    assert.strictEqual(issueAfterEvent("ngày 05 tháng 12 năm 2026", ["Căn cứ Nghị quyết ngày 20/11/2026;", "Thời gian: 8 giờ 30 ngày 20/12/2026"]), null);
  });
  it("ngày ban hành sau ngày họp: báo 'sau'", () => {
    assert.match(issueAfterEvent("ngày 12 tháng 12 năm 2026", ["Vào lúc 8h30 ngày 10/12/2026 tại hội trường"]) || "", /sau ngày họp/);
  });
  it("hợp đồng đặt cọc của chủ đầu tư cọc 10% (dots-3-note): báo giới hạn 5%; cọc 5% hoặc hợp đồng giữa hai cá nhân không báo", () => {
    const body = ["## BÊN BÁN (BÊN A): CÔNG TY CỔ PHẦN ĐẦU TƯ ANH QUÂN", "Căn nhà là nhà ở hình thành trong tương lai.", "3.1. BÊN B đặt cọc cho BÊN A số tiền là 620 triệu đồng, tương đương 10% giá bán Căn nhà."];
    assert.ok(legalThresholds(body).some((c) => /5% giá bán/.test(c)), String(legalThresholds(body)));
    assert.deepStrictEqual(legalThresholds(body.map((l) => l.replace("10%", "5%"))), []);
    assert.deepStrictEqual(legalThresholds(["Bên B đặt cọc 20% giá chuyển nhượng thửa đất cho ông Nguyễn Văn A."]), []);
  });
  it("HĐLĐ kỹ sư thử việc 90 ngày, lương thử việc 70% (dots-3-note): báo cả hai; 60 ngày, 85% không báo", () => {
    const body = ["Ông Lê Minh Tuấn, kỹ sư xây dựng.", "3.1. Trước khi ký hợp đồng, ông Lê Minh Tuấn thử việc 90 ngày.", "3.2. Lương thử việc bằng 70% mức lương chính thức."];
    const c = legalThresholds(body);
    assert.ok(c.some((x) => /Thử việc 90 ngày.*60 ngày/.test(x)) && c.some((x) => /Lương thử việc 70%.*85%/.test(x)), c.join(" | "));
    assert.deepStrictEqual(legalThresholds(body.map((l) => l.replace("90", "60").replace("70%", "85%"))), []);
    assert.ok(legalThresholds(["Hợp đồng lao động xác định thời hạn 48 tháng."]).some((x) => /36 tháng/.test(x)));
  });
  it("bên B là cá nhân (người lao động, người mua): phần ký 'BÊN B', không chức vụ, không đóng dấu; bên A công ty giữ nguyên", () => {
    const p = planAdminDocument({
      doc_type: "hop_dong",
      title: "HỢP ĐỒNG LAO ĐỘNG",
      body: "## BÊN A (BÊN SỬ DỤNG LAO ĐỘNG): CÔNG TY CỔ PHẦN ĐẦU TƯ ANH QUÂN\nĐại diện: Nguyễn Anh Quân\n## BÊN B (LAO ĐỘNG): LÊ MINH TUẤN\n## Điều 1. Công việc",
      left_signer_title: "TỔNG GIÁM ĐỐC",
      left_signer_name: "Nguyễn Anh Quân",
      signer_title: "KỸ SƯ XÂY DỰNG",
      signer_name: "Lê Minh Tuấn",
    });
    const sig = (p.blocks.filter((b) => b.kind === "table") as TableBlock[]).pop()!.rows[0];
    const [l, r] = sig.map((col) => col.map((x) => x.text).filter(Boolean));
    assert.deepStrictEqual(l, ["ĐẠI DIỆN BÊN A", "TỔNG GIÁM ĐỐC", "(Ký, ghi rõ họ tên, đóng dấu)", "Nguyễn Anh Quân"]);
    assert.deepStrictEqual(r, ["BÊN B", "(Ký, ghi rõ họ tên)", "Lê Minh Tuấn"]);
    assert.ok(!(p.blocks as ParaBlock[]).some((b) => b.text === "KỸ SƯ XÂY DỰNG"));
  });
  it("nhận diện bên: 'Bà', tên không có chữ tổ chức là cá nhân; công ty, UBND, chưa điền là tổ chức", () => {
    assert.strictEqual(partyIsPerson(["**BÊN MUA (BÊN B): BÀ PHẠM THỊ LAN**"], "B"), true);
    assert.strictEqual(partyIsPerson(["Bên B: Ông Trần Văn Bình"], "B"), true);
    assert.strictEqual(partyIsPerson(["## BÊN BÁN (BÊN A): CÔNG TY CỔ PHẦN ĐẦU TƯ ANH QUÂN"], "A"), false);
    assert.strictEqual(partyIsPerson(["## BÊN GIAO THẦU (BÊN A): Ủy ban nhân dân phường An Phú"], "A"), false);
    assert.strictEqual(partyIsPerson(["## BÊN B (BÊN NHẬN THẦU): …"], "B"), false);
    assert.strictEqual(partyIsPerson(["Không có dòng giới thiệu bên"], "B"), false);
    // Dòng của bên A không bị lấy nhầm cho bên B
    assert.strictEqual(partyIsPerson(["## BÊN A: Bà Nguyễn Thị Hoa", "## BÊN B: CÔNG TY TNHH XÂY DỰNG 3"], "B"), false);
  });
});


describe("4.11.11 · Giấy ủy quyền và dấu * dính chữ (vòng thử 3)", () => {
  const P = { org: "Công ty Cổ phần Đầu tư Anh Quân", body: "Tổng Giám đốc Nguyễn Anh Quân*uỷ*quyền cho ông Trần Văn Bình.\nThời hạn*uỷ*quyền: đến hết ngày 31/12/2027.", signer_title: "TỔNG GIÁM ĐỐC", signer_name: "Nguyễn Anh Quân" };
  const texts = (p: ReturnType<typeof planAdminDocument>) =>
    p.blocks.flatMap((b) => (b.kind === "table" ? b.rows.flatMap(([l, r]) => [...l, ...r]).map((x) => x.text) : [b.text]));
  it("công văn 'V/v Uỷ quyền …', Kính gửi để trống (dots-3-note): dựng thành GIẤY ỦY QUYỀN, không Kính gửi, hai bên ký", () => {
    const p = planAdminDocument({ ...P, doc_type: "cong_van", subject: "V/v Uỷ quyền cho ông Trần Văn Bình ký hợp đồng mua bán nhà ở" });
    const t = texts(p);
    assert.strictEqual(p.type, "giay_uy_quyen");
    assert.ok(t.includes("GIẤY ỦY QUYỀN"), t.join(" | "));
    assert.ok(!t.some((x) => /^Kính gửi/.test(x)) && !t.some((x) => /^V\/v/.test(x)), t.join(" | "));
    assert.ok(t.includes("NGƯỜI NHẬN ỦY QUYỀN") && t.includes("TỔNG GIÁM ĐỐC"), t.join(" | "));
  });
  it("doc_type giay_uy_quyen nhận trực tiếp; dấu * dính chữ thành khoảng trắng", () => {
    const p = planAdminDocument({ ...P, doc_type: "giay_uy_quyen", left_signer_name: "Trần Văn Bình" });
    const t = texts(p);
    assert.ok(t.some((x) => x.startsWith("Tổng Giám đốc Nguyễn Anh Quân uỷ quyền cho")), t.join(" | "));
    assert.ok(!t.some((x) => x.includes("*")), t.join(" | "));
    assert.ok(t.includes("Trần Văn Bình"));
  });
  it("công văn thường không bị đổi loại; '*' trong '**Điều 1.**' và số 5*3 không bị đụng", () => {
    assert.strictEqual(planAdminDocument({ ...P, doc_type: "cong_van", subject: "Đề nghị hướng dẫn thủ tục", body: "Nội dung." }).type, "cong_van");
    const p = planAdminDocument({ ...P, doc_type: "quyet_dinh", body: "**Điều 1.** Bổ nhiệm ông A.\nDiện tích 5*3 m." });
    assert.ok(texts(p).some((x) => x.includes("5*3")), texts(p).join(" | "));
  });
});

describe("4.11.11 · Soát nhất quán có thêm kiểm sự kiện (vòng thử 3, sửa công văn có sẵn)", () => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { checkConsistency } = require("../../src/core/consistency");
  const rows = [
    "Số: 118/2026/CV-AQ",
    "V/v đề nghi hướng dẫn thủ tục bàn giao nhà ở cho khách hàng",
    "Căn cứ Luật Kinh doanh bất động sản năm 2014 và Nghị định số 99/2015/NĐ-CP quy định chi tiết Luật Nhà ở;",
    "Công ty là chủ đầu tư dư án Khu nhà ở. Dự kiến bàn giao vào Thứ Hai, ngày 15/12/2026.",
    "Trân trong cảm ơn./.",
  ].map((text, i) => ({ i, text }));
  it("bắt thứ sai (15/12/2026 là Thứ Ba), luật và nghị định hết hiệu lực, 'đề nghi', 'dư án', 'Trân trong'", () => {
    const f = checkConsistency(rows, { checks: ["fact"] }).issues.map((x: { p: number; message: string }) => `${x.p}: ${x.message}`);
    const all = f.join("\n");
    assert.match(all, /3: Thứ trong tuần sai: 15\/12\/2026 là Thứ Ba/);
    assert.match(all, /2: Dẫn văn bản ĐÃ HẾT HIỆU LỰC: .*Kinh doanh bất động sản 2014.*99\/2015/);
    assert.match(all, /“đề nghi” → “đề nghị”/);
    assert.match(all, /“dư án” → “dự án”/);
    assert.match(all, /“Trân trong” → “trân trọng”/);
  });
  it("văn bản sạch: không báo; mặc định (không nêu checks) có chạy kiểm sự kiện", () => {
    const ok = [{ i: 0, text: "Kính đề nghị Sở hướng dẫn thủ tục. Họp vào Thứ Ba, ngày 15/12/2026 theo Luật Kinh doanh bất động sản năm 2023." }];
    assert.deepStrictEqual(checkConsistency(ok, { checks: ["fact"] }).issues, []);
    assert.ok(checkConsistency(rows).issues.some((x: { kind: string }) => x.kind === "fact"));
  });
});

describe("4.11.11 · Kỹ năng tự kích hoạt theo ranh giới từ (vòng thử 3)", () => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { BUILTIN_SKILLS, autoSelectSkills } = require("../../src/core/skillpack");
  const names = (m: string) => autoSelectSkills(m, BUILTIN_SKILLS.filter((s: { hosts?: string[] }) => !s.hosts || s.hosts.includes("Word"))).map((s: { name: string }) => s.name);
  it("'chỗ ghi thứ không khớp với ngày' không kéo kỹ năng chuẩn bị họp; 'rà soát văn bản' kéo kỹ năng văn bản hành chính", () => {
    const n = names("Rà soát văn bản đang mở và sửa trực tiếp: lỗi chính tả, căn cứ đã hết hiệu lực, chỗ ghi thứ không khớp với ngày.");
    assert.ok(!n.includes("chuan-bi-hop"), n.join(", "));
    assert.ok(n.includes("soan-thao-van-ban-hanh-chinh"), n.join(", "));
  });
  it("từ khóa vẫn khớp như trước khi đứng đúng chỗ: 'chuẩn bị họp với Sở', 'soạn hợp đồng lao động', 'giấy ủy quyền'", () => {
    assert.ok(names("Chuẩn bị họp với Sở Xây dựng ngày mai").includes("chuan-bi-hop"));
    assert.ok(names("Soạn hợp đồng lao động cho kỹ sư").includes("soan-hop-dong"));
    assert.ok(names("Soạn giấy ủy quyền ký hợp đồng mua bán").includes("soan-thao-van-ban-hanh-chinh"));
  });
});

describe("4.11.11 · Loại văn bản model chọn nhầm, ví dụ kỹ năng chép nguyên (vòng thử 3, Cohere)", () => {
  it("giay_moi kèm title 'GIẤY ỦY QUYỀN' (Cohere): dựng thành giấy ủy quyền", () => {
    const p = planAdminDocument({ doc_type: "giay_moi", title: "GIẤY ỦY QUYỀN", org: ORG, body: "Ông Nguyễn Anh Quân ủy quyền cho ông Trần Văn Bình.", signer_title: "TỔNG GIÁM ĐỐC", signer_name: "Nguyễn Anh Quân" });
    assert.strictEqual(p.type, "giay_uy_quyen");
  });
  it("hợp đồng đặt cọc có '## BÊN GIAO THẦU (BÊN A):' không tên: báo cả thiếu tên lẫn vai trò chép từ ví dụ", () => {
    const p = planAdminDocument({ doc_type: "hop_dong", title: "HỢP ĐỒNG ĐẶT CỌC", body: "## BÊN GIAO THẦU (BÊN A):\n## BÊN NHẬN THẦU (BÊN B):\n## Điều 1. Đối tượng" });
    assert.ok(p.checks?.some((c) => /chưa có tên/.test(c)) && p.checks?.some((c) => /giao thầu \/ nhận thầu/.test(c)), String(p.checks));
    const ok = planAdminDocument({ doc_type: "hop_dong", title: "HỢP ĐỒNG THI CÔNG XÂY DỰNG", body: "## BÊN GIAO THẦU (BÊN A): CÔNG TY CỔ PHẦN ĐẦU TƯ ANH QUÂN\n## BÊN NHẬN THẦU (BÊN B): …\n## Điều 1. Đối tượng" });
    assert.ok(!ok.checks?.some((c) => /chưa có tên|giao thầu \/ nhận thầu/.test(c)), String(ok.checks));
  });
});

describe("4.11.11 · AI kể đã sửa “A” → “B” mà không ghi (vòng thử 3, Cohere)", () => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { unbackedEditClaims } = require("../../src/core/claims");
  const cohere = 'Đã sửa các lỗi chính tả:\n• Đoạn 2: đổi "đề nghi" → "đề nghị"\n• Đoạn 5: sửa "dư án" → "dự án"\n• Đoạn 5: chỉnh "Thứ" thành "Thứ Hai"\n• Đoạn 7: sửa "Trân trong" → "Trân trọng"';
  const wrote = [JSON.stringify({ find: "đề nghi", replace: "đề nghị", paragraph_index: 2 }), JSON.stringify({ find: "dư án", replace: "dự án" })];
  it("bắt các chỗ kể mà không ghi; không kể chỗ đã ghi", () => {
    const w = unbackedEditClaims(cohere, wrote);
    assert.ok(w && /Trân trọng/.test(w.text) && /Thứ Hai/.test(w.text), w && w.text);
    assert.ok(!w.text.includes("“đề nghi” → “đề nghị”") && !w.text.includes("“dư án”"), w.text);
  });
  it("mọi chỗ kể đều có lệnh ghi (dạng set_paragraph_text cả đoạn): không báo; lượt không ghi gì: để cảnh báo cũ lo", () => {
    const all = [JSON.stringify({ index: 7, text: "Trân trọng cảm ơn./." }), JSON.stringify({ index: 5, text: "… vào Thứ Hai, ngày 15/12/2026." }), ...wrote];
    assert.strictEqual(unbackedEditClaims(cohere, all), null);
    assert.strictEqual(unbackedEditClaims(cohere, []), null);
  });
  it("bảng kể sửa không có mũi tên (Nemotron) và câu trả lời không kể chỗ sửa: không báo", () => {
    assert.strictEqual(unbackedEditClaims("2\t„đề nghi“\tđề nghị\tThiếu dấu", wrote), null);
    assert.strictEqual(unbackedEditClaims("Đã soạn xong giấy mời.", wrote), null);
  });
});

describe("4.11.11 · 'Uy ban' mất dấu (vòng thử 3, dots-3-note)", () => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { accentTypos } = require("../../src/core/adminDoc");
  it("'Uy ban nhân dân' báo; 'Ủy ban nhân dân' và kiểu cũ 'Uỷ ban nhân dân' không báo", () => {
    assert.ok(accentTypos(["Kính gửi: Uy ban nhân dân phường An Phú"]).some((x: string) => /Uy ban nhân dân/.test(x)));
    assert.deepStrictEqual(accentTypos(["Kính gửi: Ủy ban nhân dân phường An Phú"]), []);
    assert.deepStrictEqual(accentTypos(["Kính gửi: Uỷ ban nhân dân phường An Phú"]), []);
  });
});

describe("4.11.11 · Ca dễ báo nhầm theo soát mã độc lập (không được báo)", () => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { weekdayIssues, stripModelMarkup, strangeWords } = require("../../src/core/textcheck");
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { accentTypos } = require("../../src/core/adminDoc");
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { unbackedEditClaims, toolFailureClaim } = require("../../src/core/claims");
  it("#1 'Lương thử việc: 100% mức lương' không thành '00%'", () => {
    assert.deepStrictEqual(legalThresholds(["Kỹ sư xây dựng.", "Lương thử việc: 100% mức lương chính thức."]), []);
  });
  it("#2 số thứ tự không phải thứ trong tuần; ngày kèm thứ trong ngoặc vẫn kiểm", () => {
    assert.deepStrictEqual(weekdayIssues(["Phiên họp thứ 2, ngày 15/12/2026 tại hội trường."]), []);
    assert.deepStrictEqual(weekdayIssues(["Đại hội lần thứ hai, ngày 17/12/2026."]), []);
    assert.strictEqual(weekdayIssues(["Ngày 15/12/2026 (thứ 2)"]).length, 1);
    assert.strictEqual(weekdayIssues(["vào Thứ Hai, ngày 15/12/2026"]).length, 1);
  });
  it("#3 'Bên A có trách nhiệm: …' ở điều khoản, tên tổ chức nước ngoài, Kho bạc: không phải bên cá nhân", () => {
    assert.strictEqual(partyIsPerson(["## BÊN GIAO THẦU (BÊN A)", "Công ty X", "## Điều 5. Trách nhiệm", "Bên A có trách nhiệm: thanh toán đúng hạn."], "A"), false);
    assert.strictEqual(partyIsPerson(["## BÊN B: SAMSUNG ELECTRONICS VIỆT NAM"], "B"), false);
    assert.strictEqual(partyIsPerson(["## BÊN A: Kho bạc Nhà nước tỉnh Bình Minh"], "A"), false);
    assert.strictEqual(partyIsPerson(["## BÊN B (LAO ĐỘNG): LÊ MINH TUẤN"], "B"), true);
  });
  it("#4 phạt cọc, cọc giữa hai cá nhân không bị báo trần 5%", () => {
    const org = "## BÊN BÁN (BÊN A): CÔNG TY CỔ PHẦN ĐẦU TƯ ANH QUÂN";
    assert.deepStrictEqual(legalThresholds([org, "Nhà ở hình thành trong tương lai, chủ đầu tư Anh Quân.", "Bên vi phạm chịu phạt cọc 50% giá bán."]), []);
    assert.deepStrictEqual(legalThresholds(["## BÊN A: Ông Nguyễn Văn A", "## BÊN B: Bà Trần Thị B", "Căn hộ do chủ đầu tư X xây dựng.", "Bên B đặt cọc 10% giá bán."]), []);
  });
  it("#5 phụ cấp đứng sau lương thử việc 85% không bị đọc là lương thử việc", () => {
    assert.deepStrictEqual(legalThresholds(["Kỹ sư.", "Lương thử việc hưởng 85% mức lương và phụ cấp ăn trưa 30%."]), []);
  });
  it("#6 dòng căn cứ và số văn bản không bị coi là giờ họp", () => {
    assert.strictEqual(issueAfterEvent("ngày 05 tháng 12 năm 2026", ["Căn cứ Nghị quyết số 12 Hội đồng quản trị ngày 01/12/2026;", "Thời gian: 8h30 ngày 20/12/2026"]), null);
  });
  it("#7 chữ đậm hợp lệ giữ nguyên chữ; dấu * dính chữ sau dấu phẩy vẫn gỡ", () => {
    assert.strictEqual(stripModelMarkup("Đại diện: **Ông A**, Chức vụ: **Giám đốc**"), "Đại diện: **Ông A**, Chức vụ: **Giám đốc**");
    assert.strictEqual(stripModelMarkup('gọi là "Bên uỷ quyền"),*uỷ*quyền cho'), 'gọi là "Bên uỷ quyền"), uỷ quyền cho');
  });
  it("#8 cụm đúng chính tả không bị báo sai dấu", () => {
    assert.deepStrictEqual(accentTypos(["Việc này có phần chậm.", "Ông Trần Trọng Nghĩa có mặt.", "Nhóm cổ động viên.", "Kế hoạch tan thành mây khói."]), []);
  });
  it("#9 email, trang web không phải từ lạ", () => {
    assert.deepStrictEqual(strangeWords(["Liên hệ: anhquan@gmail.com, www.congtyanhquan.vn"]), []);
  });
  it("#10 chép lại danh sách đề xuất kèm 'chưa sửa', hoặc trích có dấu …: không cảnh báo", () => {
    assert.strictEqual(unbackedEditClaims('Các lỗi tôi chưa sửa, đề xuất: "Trân trong" → "Trân trọng".', ['{"x":1}']), null);
    assert.strictEqual(unbackedEditClaims('Đã sửa "dư án…" → "dự án…".', ['{"replace":"dự án"}']), null);
  });
  it("#11 công văn có Kính gửi về việc ủy quyền giao dịch ngân hàng vẫn là công văn", () => {
    assert.strictEqual(planAdminDocument({ doc_type: "cong_van", org: ORG, subject: "V/v ủy quyền giao dịch tại Ngân hàng", to: "Ngân hàng TMCP Phát Triển", body: "Nội dung." }).type, "cong_van");
  });
  it("#12 hợp đồng thuê nhắc hợp đồng lao động, thời hạn thuê 60 tháng: không báo trần 36 tháng", () => {
    assert.deepStrictEqual(legalThresholds(["Bên thuê xuất trình hợp đồng lao động để chứng minh thu nhập.", "Thời hạn thuê: 60 tháng."]), []);
  });
  it("#13 câu giả định / phủ định về công cụ không bị coi là đổ lỗi", () => {
    const f = (answer: string) => toolFailureClaim({ userText: "Soạn công văn", answer, wrote: false, toolsCalled: [] });
    assert.strictEqual(f("Nếu công cụ gặp lỗi, tôi sẽ báo lại."), null);
    assert.strictEqual(f("Công cụ không gặp lỗi gì."), null);
  });
});

describe("4.11.11 · Kiểm từ lạ trong chữ model vừa ghi khi sửa văn bản (vòng thử 3, dots-3-note)", () => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { newTextCheck } = require("../../src/core/textcheck");
  it("bắt 'DeniedĐến nay', 'Collaboratively配合', 'bringing to'", () => {
    assert.match(newTextCheck("Công ty là chủ đầu tư dự án. DeniedĐến nay công ty đã hoàn thành 48 căn.", "Công ty là chủ đầu tư dư án") || "", /Denied/);
    assert.match(newTextCheck("b) Collaboratively配合 trong việc bàn giao", "") || "", /Collaboratively|配合/);
    assert.match(newTextCheck("giải quyết bằng bringing to tòa án", "") || "", /bringing/);
  });
  it("không báo: tên riêng tiếng Anh đứng riêng, từ lạ đã có sẵn trong chữ cũ, chữ tiếng Việt thường", () => {
    assert.strictEqual(newTextCheck("Dự án Khu nhà ở Anh Quân Riverside tại phường An Phú.", ""), null);
    assert.strictEqual(newTextCheck("Trân trọng cảm ơn./.", "Trân trong cảm ơn./."), null);
    assert.strictEqual(newTextCheck("Phần mềm website của công ty", "Phần mềm website"), null);
  });
});

describe("4.11.11 · Giấy ủy quyền để trống trích yếu (bản cuối, dots-3-note)", () => {
  it("check_now nhắc viết trích yếu", () => {
    const p = planAdminDocument({ doc_type: "giay_uy_quyen", org: ORG, body: "Ông A ủy quyền cho ông B.", signer_title: "TỔNG GIÁM ĐỐC", signer_name: "Nguyễn Anh Quân" });
    assert.ok(p.checks?.some((c) => /Trích yếu/.test(c)), String(p.checks));
  });
});
