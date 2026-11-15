/** Đợt 1B — Word soát và tra: soát nhất quán bằng mã, thể thức trang, so phiên bản.
 *  Dữ liệu mẫu test/fixtures/dot1b-doan.json do `node kiem-thu/soat-thu.mjs json test/fixtures/dot1b-doan.json` sinh ra
 *  (cùng nội dung với các tệp .docx thử dùng trong Word): bản sạch, bản cài sẵn 9 lỗi, phiên bản khác lệch 10 chỗ. */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import { ParaRow, buildOutline } from "../../src/core/outline";
import { checkAmounts, checkBlanks, checkConsistency, checkNumbering, checkReferences, checkTerms, findDefinitions, parseFigure, parseVnNumber } from "../../src/core/consistency";
import { VAN_BAN_HANH_CHINH, PT_PER_MM, checkPageFormat, customStandard, ptToMm } from "../../src/core/pageformat";
import { compareVersions, describeChanges, lcsPairs, similarity } from "../../src/core/compare";

interface Doan {
  text: string;
  style: string;
  key: string;
}
const fx: { doan: Record<string, Doan[]>; loi: { loai: string; o: string }[]; v2: { dieu: string; kieu?: string }[] } = JSON.parse(
  fs.readFileSync(path.join(__dirname, "../fixtures/dot1b-doan.json"), "utf8")
);

/** Dựng các đoạn như Word trả về: kiểu Heading 1–3 mang số do Word tự đánh ("Điều 1." / "1.1." / "a)"), còn lại là chữ thường */
function toRows(list: Doan[]): ParaRow[] {
  let a = 0;
  let k = 0;
  let j = 0;
  return list.map((d, i) => {
    const row: ParaRow = { i, text: d.text, style: d.style, outlineLevel: 10 };
    if (d.style === "Heading1") {
      a++;
      k = 0;
      j = 0;
      Object.assign(row, { listString: `Điều ${a}.`, listLevel: 0, listId: 2, outlineLevel: 1 });
    } else if (d.style === "Heading2") {
      k++;
      j = 0;
      Object.assign(row, { listString: `${a}.${k}.`, listLevel: 1, listId: 2, outlineLevel: 2 });
    } else if (d.style === "Heading3") {
      j++;
      Object.assign(row, { listString: `${"abcdefgh"[j - 1]})`, listLevel: 2, listId: 2, outlineLevel: 3 });
    }
    return row;
  });
}
const rows = (list: (Partial<ParaRow> & { text: string })[]): ParaRow[] => list.map((r, i) => ({ i, ...r }));
const typed = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
const kinds = (rs: ParaRow[]) => checkConsistency(rs).issues.map((x) => x.kind);

describe("Đợt 1B – soát nhất quán: bài nghiệm thu (9 lỗi cài sẵn, bản sạch không báo nhầm)", () => {
  it("bản cài lỗi: tìm đủ 9 — 3 thuật ngữ lệch, 2 dẫn chiếu hỏng, 1 chỗ nhảy số, 1 số tiền lệch chữ, 2 chỗ trống", () => {
    const rs = toRows(fx.doan.loi);
    const r = checkConsistency(rs);
    const count = (k: string) => r.issues.filter((x) => x.kind === k).length;
    assert.deepStrictEqual(
      { term: count("term"), reference: count("reference"), numbering: count("numbering"), amount: count("amount"), blank: count("blank"), party: count("party") },
      { term: 3, reference: 2, numbering: 1, amount: 1, blank: 2, party: 0 }
    );
    assert.strictEqual(r.issues.length, 9);
    assert.strictEqual(fx.loi.length, 9);
    // Đúng chỗ đã cài: mỗi lỗi cài sẵn nằm ở một đoạn có điểm được báo
    const at = (key: string) => fx.doan.loi.findIndex((d) => d.key === key);
    for (const e of fx.loi) {
      const p = at(e.o);
      assert.ok(r.issues.some((x) => x.p === p || (x.also || []).includes(p)), `chưa báo lỗi cài ở ${e.o} (${e.loai})`);
    }
  });

  it("bản cài lỗi: thông báo nêu đúng số điều và nội dung", () => {
    const r = checkConsistency(toRows(fx.doan.loi));
    const msg = (k: string) => r.issues.filter((x) => x.kind === k).map((x) => `${x.clause}|${x.message}`);
    assert.ok(msg("numbering")[0].includes("sau 2.2 là 2.4"));
    assert.ok(msg("reference").some((m) => m.startsWith("11.2|") && m.includes("khoản 6 Điều 7") && m.includes("Điều 7")));
    assert.ok(msg("reference").some((m) => m.includes("“Điều 19”")));
    assert.ok(msg("term").some((m) => m.startsWith("5.2|") && m.includes("“Ngày Bàn Giao”") && m.includes("“ngày bàn giao”")));
    assert.ok(msg("term").some((m) => m.startsWith("6.3(a)|") && m.includes("“Tiền thuê”")));
    assert.ok(msg("term").some((m) => m.startsWith("10.2|") && m.includes("“Phí Quản Lý”") && m.includes("“Phí Dịch Vụ”")));
    assert.ok(msg("amount")[0].startsWith("4.1|") && msg("amount")[0].includes("125.000.000"));
    assert.ok(msg("blank").some((m) => m.startsWith("7.1|")));
  });

  it("bản sạch và phiên bản khác: không báo điểm nào", () => {
    for (const v of ["sach", "v2"]) {
      const r = checkConsistency(toRows(fx.doan[v]));
      assert.deepStrictEqual(
        r.issues.map((x) => `${x.kind} ¶${x.p}: ${x.message}`),
        [],
        `bản ${v} bị báo nhầm`
      );
    }
  });

  it("nhận đủ thuật ngữ định nghĩa và tên hai bên của văn bản thử", () => {
    const r = checkConsistency(toRows(fx.doan.sach));
    assert.deepStrictEqual(r.parties.sort(), ["Bên Cho Thuê", "Bên Thuê"]);
    for (const t of ["Tòa Nhà", "Diện Tích Thuê", "Ngày Bàn Giao", "Tiền Thuê", "Phí Dịch Vụ", "Tiền Đặt Cọc", "Ngày Làm Việc"]) assert.ok(r.terms.includes(t), t);
  });

  it("chỉ chạy loại soát được yêu cầu", () => {
    const r = checkConsistency(toRows(fx.doan.loi), { checks: ["blank", "amount"] });
    assert.deepStrictEqual(Array.from(new Set(r.issues.map((x) => x.kind))).sort(), ["amount", "blank"]);
  });
});

describe("Đợt 1B – soát nhất quán: đánh số", () => {
  it("số gõ tay nhảy, trùng; danh sách đánh số lại từ đầu thì không báo", () => {
    const rs = typed(["Điều 1. Phạm vi", "1.1 Một.", "1.2 Hai.", "1.2 Hai lần nữa.", "1.5 Khoản thứ năm.", "Điều 2. Giá", "2.1 Một.", "Điều 4. Hiệu lực", "4.1 Một."]);
    const out = checkNumbering(rs, buildOutline(rs)).map((x) => x.message);
    assert.strictEqual(out.length, 3);
    assert.ok(out[0].includes("trùng") && out[0].includes("1.2"));
    assert.ok(out[1].includes("sau 1.2 là 1.5") && out[1].includes("thiếu 2 số"));
    assert.ok(out[2].includes("sau Điều 2 là Điều 4"));
  });

  it("điểm a, b, c, d, đ, e, g theo bảng chữ cái tiếng Việt không phải nhảy số; a, b, d là nhảy", () => {
    const vn = typed(["Điều 1. Quyền", "1.1 Các quyền:", "a) một;", "b) hai;", "c) ba;", "d) bốn;", "đ) năm;", "e) sáu;", "g) bảy;", "h) tám."]);
    assert.deepStrictEqual(checkNumbering(vn, buildOutline(vn)), []);
    const en = typed(["Điều 1. Quyền", "1.1 Các quyền:", "a) một;", "b) hai;", "c) ba;", "d) bốn;", "e) năm;", "f) sáu."]);
    assert.deepStrictEqual(checkNumbering(en, buildOutline(en)), []);
    const jump = typed(["Điều 1. Quyền", "1.1 Các quyền:", "a) một;", "b) hai;", "d) bốn."]);
    const out = checkNumbering(jump, buildOutline(jump));
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].message.includes("nhảy"));
  });

  it("khoản bị xóa có theo dõi làm số nhảy: báo kèm lý do", () => {
    const rs = rows([{ text: "Điều 1. Phạm vi" }, { text: "1.1 Một." }, { text: "", deleted: true }, { text: "1.3 Ba." }]);
    const out = checkNumbering(rs, buildOutline(rs));
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].message.includes("đã bị xóa có theo dõi"));
  });

  it("khoản đầu tiên của điều không mang số 1", () => {
    const rs = typed(["Điều 1. Phạm vi", "1.1 Một.", "Điều 2. Giá", "2.2 Hai."]);
    const out = checkNumbering(rs, buildOutline(rs));
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].message.includes("không có 2.1"));
  });

  it("các điều đánh số liên tục qua nhiều chương: không báo; phụ lục đánh số lại từ 1: không báo", () => {
    const rs = typed(["Chương I. Quy định chung", "Điều 1. Phạm vi", "Điều 2. Đối tượng", "Chương II. Quy định cụ thể", "Điều 3. Giá", "Điều 4. Thanh toán", "PHỤ LỤC 01. BẢNG GIÁ", "1. Giá thuê", "2. Phí"]);
    assert.deepStrictEqual(checkNumbering(rs, buildOutline(rs)), []);
  });
});

describe("Đợt 1B – soát nhất quán: dẫn chiếu", () => {
  const base = ["Điều 1. Phạm vi", "1.1 Một.", "1.2 Hai.", "Điều 2. Giá", "2.1 Giá là 100.", "Điều 3. Hiệu lực"];
  const refs = (extra: string[]) => {
    const rs = typed([...base, ...extra]);
    return checkReferences(rs, buildOutline(rs)).map((x) => x.message);
  };

  it("dẫn chiếu đúng thì không báo: Điều, khoản … Điều …, Điều x.y, khoản … Điều này", () => {
    assert.deepStrictEqual(refs(["3.1 Theo Điều 2, khoản 2 Điều 1 và Điều 1.1 của hợp đồng này.", "3.2 Như nêu tại khoản 1 Điều này."]), []);
  });

  it("dẫn chiếu tới điều, khoản, điểm không có", () => {
    const out = refs(["3.1 Theo Điều 9 và khoản 5 Điều 1.", "3.2 Theo điểm c khoản 2 Điều 1 và khoản 7 Điều này."]);
    assert.strictEqual(out.length, 4);
    assert.ok(out[0].includes("“Điều 9”"));
    assert.ok(out[1].includes("khoản 5 Điều 1") && out[1].includes("có Điều 1"));
    assert.ok(out[2].includes("điểm c khoản 2 Điều 1") && out[2].includes("1.2"));
    assert.ok(out[3].includes("khoản 7 Điều này") && out[3].includes("Điều 3"));
  });

  it("dẫn chiếu tới văn bản pháp luật và văn bản khác: không soát", () => {
    assert.deepStrictEqual(
      refs([
        "3.1 Theo Điều 328 của Bộ luật Dân sự và khoản 3 Điều 40 Luật Kinh doanh bất động sản.",
        "3.2 Theo quy định tại Điều 12, Điều 13 và Điều 14 của Nghị định về quản lý dự án.",
        "3.3 Như đã thỏa thuận tại Điều 7 của Hợp đồng số 05/2025/HĐMB và Điều 9 Thỏa thuận khung ngày 01/01/2025.",
        "3.4 Luật Đất đai (Điều 45) quy định khác.",
      ]),
      []
    );
  });

  it("câu mở đầu bằng dẫn chiếu vẫn được soát; dòng tiêu đề của chính điều đó thì không", () => {
    const out = refs(["Điều 8 của hợp đồng này và Điều 2 vẫn có hiệu lực."]);
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].includes("“Điều 8”"));
  });

  it("cùng một dẫn chiếu hỏng ở nhiều đoạn: gộp thành một điểm", () => {
    const rs = typed([...base, "3.1 Theo Điều 9.", "3.2 Cũng theo Điều 9."]);
    const out = checkReferences(rs, buildOutline(rs));
    assert.strictEqual(out.length, 1);
    assert.strictEqual(out[0].count, 2);
    assert.deepStrictEqual(out[0].also, [7]);
  });

  it("điểm viết liền trong chữ của khoản (không tách đoạn) vẫn được coi là có", () => {
    const rs = typed(["Điều 1. Phạm vi", "1.1 Gồm: (a) nhà xưởng; (b) văn phòng.", "Điều 2. Giá", "2.1 Theo Điều 1.1(b) và Điều 1.1(d)."]);
    const out = checkReferences(rs, buildOutline(rs)).map((x) => x.message);
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].includes("1.1(d)"));
  });

  it("phụ lục: chỉ soát khi văn bản có phụ lục nằm trong tệp", () => {
    assert.deepStrictEqual(refs(["3.1 Theo Phụ lục 02 kèm theo."]), []);
    const rs = typed([...base, "3.1 Theo Phụ lục 01 và Phụ lục 03.", "PHỤ LỤC 01. BẢNG GIÁ", "1. Giá thuê", "PHỤ LỤC 02. TIẾN ĐỘ", "1. Đợt một"]);
    const out = checkReferences(rs, buildOutline(rs)).map((x) => x.message);
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].includes("Phụ lục 03") && out[0].includes("Phụ lục 01"));
  });

  it("trường dẫn chiếu hỏng của Word", () => {
    const out = refs(["3.1 Theo Điều Error! Reference source not found. của hợp đồng."]);
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].includes("Trường dẫn chiếu hỏng"));
  });

  it("văn bản không có điều khoản đánh số: không soát dẫn chiếu", () => {
    const rs = typed(["Kính gửi: Quý công ty", "Căn cứ Điều 5 của hợp đồng đã ký, chúng tôi đề nghị thanh toán."]);
    assert.deepStrictEqual(checkReferences(rs, buildOutline(rs)), []);
  });
});

describe("Đợt 1B – soát nhất quán: thuật ngữ và tên các bên", () => {
  it("các cách định nghĩa: “X” là…, sau đây gọi là “X”, (“X”)", () => {
    const d = findDefinitions(
      typed(["“Dự Án” có nghĩa là dự án khu nhà ở.", "Công ty A (sau đây gọi là “Bên Bán”) và Công ty B (“Bên Mua”).", "Khoản tiền này (gọi tắt là “Giá Bán”) được trả một lần."]).map((r) => ({ i: r.i, text: r.text }))
    );
    assert.deepStrictEqual([...d.terms.keys()].sort(), ["Bên Bán", "Bên Mua", "Dự Án", "Giá Bán"]);
    assert.deepStrictEqual([...d.parties.keys()].sort(), ["Bên Bán", "Bên Mua"]);
  });

  it("chữ thường trong chính câu định nghĩa, dòng tiêu đề và dòng viết hoa toàn bộ: không tính là lệch", () => {
    const rs = rows([
      { text: "HỢP ĐỒNG VỀ GIÁ BÁN CĂN HỘ" },
      { text: "“Giá Bán” là giá bán căn hộ ghi tại Điều 2." },
      { text: "Giá bán và thanh toán", listString: "Điều 2.", listLevel: 0, listId: 1, outlineLevel: 1 },
      { text: "Giá Bán là 3 tỷ đồng.", listString: "2.1.", listLevel: 1, listId: 1, outlineLevel: 2 },
    ]);
    assert.deepStrictEqual(checkTerms(rs, buildOutline(rs)).issues, []);
  });

  it("gộp mọi chỗ viết lệch của cùng một cách viết thành một điểm, kèm số chỗ", () => {
    const rs = typed(["“Giá Bán” là giá ghi tại hợp đồng.", "Bên mua trả giá bán một lần.", "Sau khi nhận đủ giá bán thì bàn giao.", "Phần Giá bán còn lại trả sau."]);
    const out = checkTerms(rs, buildOutline(rs)).issues;
    assert.strictEqual(out.length, 2);
    assert.ok(out.some((x) => x.message.includes("“giá bán”") && x.count === 2 && x.p === 1 && (x.also || [])[0] === 2));
    assert.ok(out.some((x) => x.message.includes("“Giá bán”") && x.count === 1));
  });

  it("tên bên không có trong định nghĩa: báo một bên có thể đang được gọi bằng nhiều tên", () => {
    const rs = typed([
      "BÊN BÁN: Công ty A, sau đây gọi là “Bên Bán”.",
      "BÊN MUA: Công ty B, sau đây gọi là “Bên Mua”.",
      "Bên Bán giao hàng cho Bên Mua trong 10 ngày.",
      "Bên B thanh toán cho Bên Bán trong 30 ngày; bên thứ ba không có quyền gì.",
      "Bên B chịu phí chuyển khoản.",
    ]);
    const out = checkTerms(rs, buildOutline(rs)).issues;
    assert.strictEqual(out.length, 1);
    assert.strictEqual(out[0].kind, "party");
    assert.ok(out[0].message.includes("“Bên B”") && out[0].message.includes("2 chỗ") && out[0].message.includes("Bên Bán, Bên Mua"));
  });

  it("một bên được định nghĩa hai tên và cả hai tên đều được dùng", () => {
    const rs = typed(["BÊN A (BÊN BÁN): Công ty A.", "BÊN B (BÊN MUA): Công ty B.", "Bên A giao hàng cho Bên B.", "Bên Bán xuất hóa đơn cho Bên B."]);
    const out = checkTerms(rs, buildOutline(rs)).issues.filter((x) => x.kind === "party");
    assert.strictEqual(out.length, 1);
    assert.ok(out[0].message.includes("“Bên A”") && out[0].message.includes("“Bên Bán”"));
  });

  it("tên bên đứng liền trước một thuật ngữ khác không bị ghép thành tên lạ", () => {
    const rs = typed(["Công ty A (“Bên Thuê”) và Công ty B (“Bên Cho Thuê”).", "“Tiền Thuê” là tiền thuê hằng tháng.", "Bên Thuê trả cho Bên Cho Thuê Tiền Thuê đúng hạn."]);
    assert.deepStrictEqual(checkTerms(rs, buildOutline(rs)).issues, []);
  });

  it("tên riêng (công ty, ngân hàng, địa danh) không bị coi là thuật ngữ chưa định nghĩa", () => {
    const rs = typed(["“Ngày Bàn Giao” là ngày giao nhà.", "Công ty TNHH Lam Giang chuyển tiền vào Ngân hàng Đông Hải tại Thành phố Hà Tĩnh trước Ngày Bàn Giao."]);
    assert.deepStrictEqual(checkTerms(rs, buildOutline(rs)).issues, []);
  });
});

describe("Đợt 1B – soát nhất quán: chỗ trống", () => {
  const blanks = (texts: string[]) => {
    const rs = typed(texts);
    return checkBlanks(rs, buildOutline(rs)).map((x) => x.p);
  };
  it("các kiểu chỗ trống thường gặp", () => {
    assert.deepStrictEqual(blanks(["Số: …/2026/HĐ", "ngày ... tháng ... năm 2026", "Địa chỉ: ____________", "Giá bán: [●] đồng", "Thời hạn: [điền số ngày] ngày", "Người đại diện: TBD", "Hôm nay, ngày    tháng    năm 2026"]), [0, 1, 2, 3, 4, 5, 6]);
  });
  it("không báo nhầm: “v.v...”, ba chấm cuối câu, dòng ký tên, dấu ngoặc có chữ", () => {
    assert.deepStrictEqual(blanks(["như kho hàng, lưu trú v.v... nếu chưa được chấp thuận", "Hai bên sẽ thỏa thuận sau…", "Ký tên: ____________", "(Ký, ghi rõ họ tên và đóng dấu) ...............", "Xem [¶12] và mục [2] của danh sách."]), []);
  });
  it("nhiều chỗ trống trong một đoạn: một điểm, có đếm", () => {
    const rs = typed(["Ông ......... sinh ngày ... tháng ... năm ......"]);
    const out = checkBlanks(rs, buildOutline(rs));
    assert.strictEqual(out.length, 1);
    assert.strictEqual(out[0].count, 4);
  });
});

describe("Đợt 1B – soát nhất quán: số bằng số và bằng chữ", () => {
  it("đọc số viết bằng chữ tiếng Việt", () => {
    const v = (s: string) => parseVnNumber(s.split(" "))?.value;
    assert.strictEqual(v("một tỷ năm trăm triệu"), 1_500_000_000);
    assert.strictEqual(v("Hai mươi lăm triệu ba trăm nghìn"), 25_300_000);
    assert.strictEqual(v("một trăm linh năm"), 105);
    assert.strictEqual(v("mười lăm"), 15);
    assert.strictEqual(v("hai mươi mốt"), 21);
    assert.strictEqual(v("ba trăm sáu mươi triệu đồng"), 360_000_000);
    assert.strictEqual(v("một nghìn hai trăm tỷ"), 1_200_000_000_000);
    assert.strictEqual(v("hai triệu rưỡi"), 2_500_000);
    assert.strictEqual(v("chín trăm chín mươi chín nghìn chín trăm chín mươi chín"), 999_999);
    assert.strictEqual(parseVnNumber("đã gồm thuế".split(" ")), null);
  });

  it("đọc số viết bằng chữ số", () => {
    assert.strictEqual(parseFigure("1.500.000.000"), 1_500_000_000);
    assert.strictEqual(parseFigure("450,000"), 450_000);
    assert.strictEqual(parseFigure("05"), 5);
    assert.strictEqual(parseFigure("1,5"), 1.5);
    assert.strictEqual(parseFigure("12/2026"), null);
  });

  const amounts = (texts: string[]) => {
    const rs = typed(texts);
    return checkAmounts(rs, buildOutline(rs));
  };
  it("khớp thì không báo: tiền, thời hạn, phần trăm, số có đơn vị triệu / tỷ, chữ trước số sau", () => {
    assert.deepStrictEqual(
      amounts([
        "Giá là 1.500.000.000 đồng (Bằng chữ: Một tỷ năm trăm triệu đồng chẵn).",
        "Trong vòng 30 (ba mươi) ngày và 05 (năm) năm, phạt 8% (tám phần trăm).",
        "Mức bảo hiểm 80 tỷ đồng (tám mươi tỷ đồng); đặt cọc 2,5 tỷ đồng (hai tỷ rưỡi).",
        "Gia hạn ba (03) tháng; lập thành bốn (04) bản; trong năm (2026) này.",
        "Thời hạn 20 năm (hai mươi năm).",
        "Số: 12/2026/HĐ (bản thử, nội dung không có thật); lãi 0,05% (không phẩy không năm phần trăm).",
      ]),
      []
    );
  });
  it("lệch thì báo, kèm hai giá trị", () => {
    const out = amounts(["Giá là 1.500.000.000 đồng (Một tỷ sáu trăm triệu đồng).", "Trong vòng 30 (ba mươi lăm) ngày.", "Gia hạn ba (05) tháng."]);
    assert.strictEqual(out.length, 3);
    assert.ok(out[0].message.includes("1.500.000.000") && out[0].message.includes("1.600.000.000"));
    assert.ok(out[1].message.includes("= 35"));
    assert.ok(out[2].message.includes("= 3"));
  });
});

describe("Đợt 1B – thể thức trang", () => {
  const A4 = { pageWidth: 210 * PT_PER_MM, pageHeight: 297 * PT_PER_MM, paperSize: "A4", orientation: "Portrait" };
  const sec = (top: number, bottom: number, left: number, right: number) => ({ section: 0, topMargin: top * PT_PER_MM, bottomMargin: bottom * PT_PER_MM, leftMargin: left * PT_PER_MM, rightMargin: right * PT_PER_MM, ...A4 });
  const tnr = (i: number, size: number, chars = 200) => ({ i, name: "Times New Roman", size, chars });

  it("đổi điểm sang mm đúng số Word trả về (56,7 điểm = 20 mm; 85,05 điểm = 30 mm)", () => {
    assert.strictEqual(ptToMm(56.7000007629395), 20);
    assert.strictEqual(ptToMm(85.0500030517578), 30);
    assert.strictEqual(ptToMm(42.5), 15);
  });

  it("văn bản đúng thể thức: không báo gì", () => {
    const r = checkPageFormat([sec(20, 20, 30, 15)], [tnr(0, 13), tnr(1, 14, 2000), tnr(2, 11, 40)]);
    assert.strictEqual(r.ok, true);
    assert.deepStrictEqual(r.issues, []);
    assert.strictEqual(r.sections[0].paper, "A4");
  });

  it("lề sai chuẩn: báo đúng lề nào sai, kèm số đo và chuẩn", () => {
    const r = checkPageFormat([sec(15, 20, 20, 15)], [tnr(0, 14, 2000)]);
    assert.deepStrictEqual(
      r.issues.map((x) => `${x.label}: ${x.actual} (chuẩn ${x.expected})`),
      ["Lề trên: 15 mm (chuẩn 20–25 mm)", "Lề trái: 20 mm (chuẩn 30–35 mm)"]
    );
  });

  it("phông lạ, cỡ chữ lệch, đoạn lẫn phông", () => {
    const r = checkPageFormat([sec(20, 20, 30, 15)], [tnr(0, 14, 2000), { i: 3, name: "Arial", size: 12, chars: 150 }, { i: 5, name: "", size: null, chars: 80 }, tnr(6, 16, 30)]);
    assert.deepStrictEqual(
      r.issues.map((x) => `${x.what}@${x.p}`),
      ["font@3", "mixed@5", "size@6"]
    );
    assert.ok(r.issues[0].actual.startsWith("Arial"));
  });

  it("cỡ chữ nội dung chính ngoài khoảng 13–14", () => {
    const r = checkPageFormat([sec(20, 20, 30, 15)], [tnr(0, 12, 3000), tnr(1, 14, 100)]);
    assert.deepStrictEqual(
      r.issues.map((x) => x.what),
      ["body_size"]
    );
  });

  it("khổ giấy khác A4 và giấy ngang", () => {
    const letter = { section: 0, topMargin: 72, bottomMargin: 72, leftMargin: 90, rightMargin: 50, pageWidth: 612, pageHeight: 792, paperSize: "Letter" };
    assert.ok(checkPageFormat([letter], []).issues.some((x) => x.what === "paper" && x.actual.includes("Letter")));
    const ngang = { ...sec(20, 20, 30, 15), pageWidth: 297 * PT_PER_MM, pageHeight: 210 * PT_PER_MM };
    assert.deepStrictEqual(
      checkPageFormat([ngang], []).issues.map((x) => x.what),
      ["orientation"]
    );
  });

  it("chuẩn do người dùng nêu thay chuẩn mặc định", () => {
    const std = customStandard(VAN_BAN_HANH_CHINH, { left_mm: 25, right_mm: [20, 25], font: "Arial", font_size: 12 });
    const r = checkPageFormat([sec(20, 20, 25, 20)], [{ i: 0, name: "Arial", size: 12, chars: 500 }], std);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(customStandard(VAN_BAN_HANH_CHINH, undefined), VAN_BAN_HANH_CHINH);
  });

  it("nhiều phần có lề khác nhau: nêu rõ phần nào", () => {
    const r = checkPageFormat([sec(20, 20, 30, 15), { ...sec(20, 20, 20, 15), section: 1 }], []);
    assert.deepStrictEqual(
      r.issues.map((x) => `${x.what}#${x.section}`),
      ["left#1"]
    );
  });
});

describe("Đợt 1B – so phiên bản", () => {
  it("bài nghiệm thu: hai phiên bản khác nhau 10 chỗ — liệt kê đủ 10 kèm số điều", () => {
    const r = compareVersions(toRows(fx.doan.sach), toRows(fx.doan.v2));
    assert.strictEqual(fx.v2.length, 10);
    assert.strictEqual(r.diffs.length, 10);
    assert.deepStrictEqual(
      r.diffs.map((d) => `${d.type}:${d.clause}`),
      ["changed:3.1", "changed:4.1", "changed:4.2", "changed:5.1", "changed:6.2", "changed:7.3", "changed:8.2", "only_other:9.3", "only_current:11.3", "changed:2.1"]
    );
    // Khoản mới 9.4 của phiên bản khác; khoản 11.4 thành 11.3 vì 11.3 bị bỏ
    const added = r.diffs.find((d) => d.type === "only_other");
    assert.strictEqual(added?.clause_other, "9.4");
    assert.ok(added?.other?.includes("bảo hiểm mọi rủi ro"));
    assert.deepStrictEqual(
      r.renumbered.map((x) => `${x.current}→${x.other}`),
      ["11.4→11.3"]
    );
    const d41 = r.diffs.find((d) => d.clause === "4.1");
    assert.ok(d41?.changes?.some((c) => c.includes("120.000.000") && c.includes("135.000.000")));
    const d82 = r.diffs.find((d) => d.clause === "8.2");
    assert.ok(d82?.changes?.[0].includes("thêm “Chi phí thẩm tra"));
  });

  it("hai bản giống hệt: không có chỗ khác; mốc chú thích và khoảng trắng thừa không tính là khác", () => {
    const a = toRows(fx.doan.sach);
    const b = toRows(fx.doan.sach).map((r) => ({ ...r, text: r.text.replace(/\.$/, ".\u0002").replace(/ /, "  ") }));
    const r = compareVersions(a, b);
    assert.strictEqual(r.identical, true);
    assert.deepStrictEqual(r.diffs, []);
  });

  it("đoạn đã bị xóa có theo dõi không tham gia so sánh", () => {
    const a = rows([{ text: "Điều 1. Phạm vi" }, { text: "1.1 Một." }, { text: "", deleted: true }, { text: "1.2 Hai." }]);
    const b = typed(["Điều 1. Phạm vi", "1.1 Một.", "1.2 Hai."]);
    assert.strictEqual(compareVersions(a, b).identical, true);
  });

  it("đoạn viết lại gần hết được coi là bỏ đoạn cũ, thêm đoạn mới", () => {
    const a = typed(["Điều 1. Phạm vi", "1.1 Bên A giao hàng tại kho của Bên B trong giờ hành chính.", "1.2 Hai."]);
    const b = typed(["Điều 1. Phạm vi", "1.1 Mọi tranh chấp được đưa ra trọng tài thương mại giải quyết.", "1.2 Hai."]);
    assert.deepStrictEqual(
      compareVersions(a, b).diffs.map((d) => d.type),
      ["only_current", "only_other"]
    );
  });

  it("lcsPairs và similarity", () => {
    assert.deepStrictEqual(lcsPairs(["a", "b", "c", "d"], ["a", "x", "c", "d"]), [
      [0, 0],
      [2, 2],
      [3, 3],
    ]);
    assert.deepStrictEqual(lcsPairs([], ["a"]), []);
    assert.ok(similarity("Thời hạn thuê là 05 năm", "Thời hạn thuê là 07 năm") > 0.7);
    assert.strictEqual(similarity("abc", ""), 0);
  });

  it("mô tả chỗ đổi: thay, thêm, bỏ", () => {
    assert.deepStrictEqual(describeChanges("mỗi năm tăng 5% so với năm trước", "mỗi năm tăng 7% so với năm trước"), ["…mỗi năm tăng “5%” → “7%”"]);
    assert.ok(describeChanges("Giao hàng tại kho.", "Giao hàng tại kho. Phí vận chuyển do bên mua chịu.")[0].includes("thêm “Phí vận chuyển do bên mua chịu.”"));
    assert.ok(describeChanges("Giao hàng tại kho trong giờ hành chính của bên bán.", "Giao hàng tại kho.")[0].includes("bỏ “"));
  });
});
