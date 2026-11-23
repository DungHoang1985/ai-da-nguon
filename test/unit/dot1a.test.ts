/** Đợt 1A — Word đọc đúng, sửa đúng: mục lục điều khoản, so khác mức từ, chống sửa nhầm đoạn, Hoàn tác. */
import * as assert from "assert";
import { ParaRow, buildOutline, clauseAt, clauseKey, findClause, parseNumber } from "../../src/core/outline";
import {
  addedKeys,
  applyHunks,
  changeKey,
  diffChanges,
  diffWords,
  liveMap,
  locateParagraph,
  matchesExpect,
  occurrenceAt,
  pickByKeys,
  pickWanted,
  planParagraphEdit,
  resolveAuthor,
  snapChange,
} from "../../src/core/worddiff";
import { isMarkdownNumberedList, packParagraphs, parseMarkdownLine, searchRows, searchRowsByWords } from "../../src/office/util";
import { findRefs } from "../../src/core/refs";
import { WORD_SPECS, WORD_SYSTEM_PROMPT } from "../../src/office/wordTools";
import { BUILTIN_SKILLS } from "../../src/core/skillpack";

const rows = (list: (Partial<ParaRow> & { text: string })[]): ParaRow[] => list.map((r, i) => ({ i, ...r }));
const nums = (rs: ParaRow[]) => buildOutline(rs).map((o) => `${o.i}:${o.num}`);

describe("Đợt 1A – mục lục điều khoản (buildOutline)", () => {
  it("số do Word tự đánh, nhiều cấp gắn kiểu tiêu đề: ghép nhãn cấp con với cấp cha", () => {
    // Đúng dữ liệu thăm dò 15/11/2026 (tasks/dot-1a-bang-chung/01)
    const rs = rows([
      { text: "HỢP ĐỒNG" },
      { text: "Phạm vi hợp đồng", listString: "Điều 1.", listLevel: 0, listId: 2, outlineLevel: 1 },
      { text: "Bên A giao cho Bên B…", listString: "1.1.", listLevel: 1, listId: 2, outlineLevel: 2 },
      { text: "Bên B cam kết…", listString: "1.2.", listLevel: 1, listId: 2, outlineLevel: 2 },
      { text: "tiến độ tổng thể;", listString: "a)", listLevel: 2, listId: 2, outlineLevel: 3 },
      { text: "tiến độ chi tiết.", listString: "b)", listLevel: 2, listId: 2, outlineLevel: 3 },
      { text: "Giá và thanh toán", listString: "Điều 2.", listLevel: 0, listId: 2, outlineLevel: 1 },
      { text: "Giá trị hợp đồng…", listString: "2.1.", listLevel: 1, listId: 2, outlineLevel: 2 },
      { text: "Đoạn thường thuộc 2.1" },
    ]);
    const o = buildOutline(rs);
    assert.deepStrictEqual(
      o.map((x) => x.num),
      ["Điều 1", "1.1", "1.2", "1.2(a)", "1.2(b)", "Điều 2", "2.1"]
    );
    assert.deepStrictEqual(
      o.map((x) => x.level),
      [0, 1, 1, 2, 2, 0, 1]
    );
    assert.strictEqual(o[0].end, 5, "Điều 1 kéo tới hết khoản con cuối");
    assert.strictEqual(o[2].end, 5, "1.2 gồm cả điểm a, b");
    assert.strictEqual(o[1].end, 2);
    assert.strictEqual(o[6].end, 8, "khoản cuối kéo tới hết văn bản");
    assert.ok(o.every((x) => x.source === "list"));
    assert.strictEqual(o[0].title, "Phạm vi hợp đồng");
  });

  it("nhiều cấp đặt trên đoạn: 1. / 1.1 / (a)", () => {
    const rs = rows([
      { text: "Quyền của Bên A", listString: "1.", listLevel: 0, listId: 3 },
      { text: "Được kiểm tra tiến độ.", listString: "1.1", listLevel: 1, listId: 3 },
      { text: "Được yêu cầu sửa chữa.", listString: "1.2", listLevel: 1, listId: 3 },
      { text: "bằng văn bản;", listString: "(a)", listLevel: 2, listId: 3 },
      { text: "Nghĩa vụ của Bên A", listString: "2.", listLevel: 0, listId: 3 },
    ]);
    assert.deepStrictEqual(nums(rs), ["0:1", "1:1.1", "2:1.2", "3:1.2(a)", "4:2"]);
  });

  it("số gõ tay: Điều 7. / 7.1 / 7.2. / a)", () => {
    const rs = rows([
      { text: "Điều 7. Thời hạn" },
      { text: "7.1 Thời hạn thuê là 24 tháng." },
      { text: "7.2. Gia hạn theo thỏa thuận." },
      { text: "a) bằng phụ lục hợp đồng;" },
      { text: "Điều 8: Bảo mật" },
    ]);
    const o = buildOutline(rs);
    assert.deepStrictEqual(
      o.map((x) => x.num),
      ["Điều 7", "7.1", "7.2", "7.2(a)", "Điều 8"]
    );
    assert.ok(o.every((x) => x.source === "typed"));
    assert.strictEqual(o[0].title, "Thời hạn");
    assert.strictEqual(o[1].title, "Thời hạn thuê là 24 tháng.");
    assert.strictEqual(o[0].end, 3);
  });

  it("kiểu văn bản Việt Nam: khoản 1, 2 và điểm a, b dưới Điều", () => {
    const rs = rows([
      { text: "Điều 5. Thanh toán" },
      { text: "1. Bên B thanh toán trong 30 ngày." },
      { text: "a) bằng chuyển khoản;" },
      { text: "b) bằng tiền mặt." },
      { text: "2. Chậm thanh toán thì chịu lãi." },
      { text: "Điều 6. Bàn giao" },
      { text: "1. Bên A bàn giao mặt bằng." },
    ]);
    assert.deepStrictEqual(nums(rs), ["0:Điều 5", "1:5.1", "2:5.1(a)", "3:5.1(b)", "4:5.2", "5:Điều 6", "6:6.1"]);
  });

  it("Chương, Mục, Điều: điều thuộc chương; chương mới đóng chương cũ", () => {
    const rs = rows([
      { text: "Chương I. QUY ĐỊNH CHUNG" },
      { text: "Điều 1. Phạm vi" },
      { text: "Điều 2. Đối tượng" },
      { text: "Chương II: QUYỀN VÀ NGHĨA VỤ" },
      { text: "Mục 1. Bên A" },
      { text: "Điều 3. Quyền của Bên A" },
    ]);
    const o = buildOutline(rs);
    assert.deepStrictEqual(
      o.map((x) => `${x.num}/${x.level}`),
      ["Chương I/0", "Điều 1/1", "Điều 2/1", "Chương II/0", "Mục 1/1", "Điều 3/2"]
    );
    assert.strictEqual(o[0].end, 2);
    assert.strictEqual(o[1].end, 1);
  });

  it("danh sách liệt kê nằm trong một điều không bị nhầm thành điều lớn (theo mã danh sách của Word)", () => {
    const rs = rows([
      { text: "Định nghĩa", listString: "1.", listLevel: 0, listId: 1 },
      { text: "Thanh toán", listString: "2.", listLevel: 0, listId: 1 },
      { text: "hồ sơ gồm hóa đơn", listString: "1.", listLevel: 0, listId: 7 },
      { text: "và biên bản nghiệm thu", listString: "2.", listLevel: 0, listId: 7 },
      { text: "Bảo hành", listString: "3.", listLevel: 0, listId: 1 },
    ]);
    assert.deepStrictEqual(nums(rs), ["0:1", "1:2", "2:2.1", "3:2.2", "4:3"]);
  });

  it("không nhận nhầm số tiền, ngày tháng, năm, chữ thường", () => {
    const rs = rows([
      { text: "1.500.000.000 đồng là giá trị hợp đồng." },
      { text: "15.11.2026 hai bên ký biên bản." },
      { text: "2026. Năm tài chính bắt đầu từ tháng 1." },
      { text: "Bên A có trách nhiệm bàn giao." },
      { text: "2.5% là mức lãi." },
      { text: "", listString: "" },
      { text: "Gạch đầu dòng", listString: "•", listLevel: 0 },
    ]);
    assert.deepStrictEqual(nums(rs), []);
  });

  it("chữ số La Mã thường và chữ cái i nối tiếp h", () => {
    const rs = rows([
      { text: "Điều 3. Phạt" },
      { text: "3.1 Các trường hợp:" },
      { text: "a) chậm tiến độ:" },
      { text: "(i) dưới 10 ngày;" },
      { text: "(ii) từ 10 ngày;" },
      { text: "b) sai chất lượng." },
    ]);
    assert.deepStrictEqual(nums(rs), ["0:Điều 3", "1:3.1", "2:3.1(a)", "3:3.1(a)(i)", "4:3.1(a)(ii)", "5:3.1(b)"]);
    const abc = "abcdefgh".split("").map((c) => ({ text: `${c}) mục ${c}` }));
    const rs2 = rows([{ text: "Điều 1. Danh mục" }, ...abc, { text: "i) mục i" }]);
    const o2 = buildOutline(rs2);
    assert.strictEqual(o2[o2.length - 1].num, "1(i)");
    assert.strictEqual(o2[o2.length - 1].kind, "letter");
  });

  it("bỏ qua đoạn trong bảng và đoạn đã bị xóa có theo dõi; tiêu đề không đánh số vẫn vào mục lục", () => {
    const rs = rows([
      { text: "PHẦN MỞ ĐẦU", outlineLevel: 1 },
      { text: "Điều 1. Một" },
      { text: "1. ô bảng", inTable: true },
      { text: "", deleted: true, listString: "Điều 9." },
      { text: "Điều 2. Hai" },
    ]);
    const o = buildOutline(rs);
    assert.deepStrictEqual(
      o.map((x) => `${x.i}:${x.num}:${x.kind}`),
      ["0::heading", "1:Điều 1:article", "4:Điều 2:article"]
    );
    assert.strictEqual(o[0].title, "PHẦN MỞ ĐẦU");
    assert.strictEqual(o[1].end, 3, "đoạn trong bảng và đoạn đã xóa vẫn thuộc phạm vi điều đứng trước");
  });

  it("câu dẫn chiếu mở đầu bằng 'Điều N', 'Phụ lục N', 'Phần N' không phải điều khoản (soát mã độc lập, C2)", () => {
    const rs = rows([
      { text: "Điều 1. Phạm vi" },
      { text: "Điều 3 của Hợp đồng này quy định rõ quyền của các bên." },
      { text: "Phụ lục 01 là bộ phận không tách rời của Hợp đồng này." },
      { text: "Phần 2 của khoản tiền được trả sau khi nghiệm thu." },
      { text: "Điều 12 Luật Nhà ở quy định về điều kiện của nhà ở tham gia giao dịch." },
      { text: "Section 3 of this Agreement applies." },
      { text: "Điều 2. Giá" },
      { text: "ĐIỀU 3 THANH TOÁN" },
      { text: "Điều 4 – Bàn giao" },
      { text: "PHỤ LỤC 01 MÔ TẢ MẶT BẰNG" },
    ]);
    const o = buildOutline(rs);
    assert.deepStrictEqual(
      o.map((x) => `${x.i}:${x.num}/${x.level}`),
      ["0:Điều 1/0", "6:Điều 2/0", "7:Điều 3/0", "8:Điều 4/0", "9:Phụ lục 01/0"]
    );
    assert.strictEqual(o[0].end, 5, "Điều 1 gồm cả các câu dẫn chiếu đứng sau nó");
  });

  it("khoản đứng trước bị xóa có theo dõi hoặc đánh nhảy số: khoản sau vẫn là anh em, không thành khoản con (T1)", () => {
    const rs = rows([{ text: "Điều 5. Thanh toán" }, { text: "1. Khoản một" }, { text: "", deleted: true }, { text: "3. Khoản ba" }, { text: "4. Khoản bốn" }]);
    assert.deepStrictEqual(nums(rs), ["0:Điều 5", "1:5.1", "3:5.3", "4:5.4"]);
    const auto = rows([
      { text: "Thanh toán", listString: "Điều 5.", listLevel: 0, listId: 1 },
      { text: "Khoản một", listString: "1.", listLevel: 0, listId: 9 },
      { text: "", deleted: true, listString: "2.", listLevel: 0, listId: 9 },
      { text: "Khoản ba", listString: "3.", listLevel: 0, listId: 9 },
    ]);
    assert.deepStrictEqual(nums(auto), ["0:Điều 5", "1:5.1", "3:5.3"]);
  });

  it("số đo đầu đoạn và danh sách chữ hoa gõ tay không thành điều khoản (L1, L2)", () => {
    const rs = rows([
      { text: "Điều 1. Giá" },
      { text: "1.5 triệu đồng mỗi tháng là mức phí quản lý." },
      { text: "3.5 tỷ đồng là giá trị bảo lãnh." },
      { text: "C. Công ty X" },
      { text: "D. Công ty Y" },
      { text: "9.4 Khoản không có điều 9 đứng trước" },
      { text: "1.1 Khoản đầu của Điều 1" },
      { text: "1.2 Khoản hai" },
    ]);
    assert.deepStrictEqual(nums(rs), ["0:Điều 1", "6:1.1", "7:1.2"]);
    const roman = rows([{ text: "I. QUY ĐỊNH CHUNG" }, { text: "Điều 1. A" }, { text: "II. QUY ĐỊNH CỤ THỂ" }, { text: "V. Lạc số" }]);
    assert.deepStrictEqual(nums(roman), ["0:I", "1:Điều 1", "2:II"]);
  });

  it("parseNumber: nhãn danh sách và số gõ tay", () => {
    assert.strictEqual(parseNumber("Điều 12.", true)?.kind, "article");
    assert.strictEqual(parseNumber("ĐIỀU 12", true)?.n, 12);
    assert.deepStrictEqual(parseNumber("12.3.", true)?.parts, [12, 3]);
    assert.strictEqual(parseNumber("(a)", true)?.letter, "a");
    assert.strictEqual(parseNumber("đ)", true)?.letter, "đ");
    assert.strictEqual(parseNumber("a.", true)?.kind, "letter");
    assert.strictEqual(parseNumber("a. Ghi chú", false), null, "'a.' gõ tay không coi là điểm");
    assert.strictEqual(parseNumber("Phụ lục 01. Bảng giá", false)?.kind, "chapter");
    assert.strictEqual(parseNumber("II. NỘI DUNG", false)?.kind, "chapter");
    assert.strictEqual(parseNumber("Điều 5.3 Thanh toán", false)?.kind, "dotted");
    assert.strictEqual(parseNumber("Theo Điều 5 thì", false), null);
  });
});

describe("Đợt 1A – tìm điều khoản theo cách gọi của người dùng", () => {
  const o = buildOutline(
    rows([
      { text: "Chương II. Thanh toán" },
      { text: "Điều 5. Thanh toán" },
      { text: "5.1 Kỳ hạn" },
      { text: "5.2 Phương thức" },
      { text: "a) chuyển khoản" },
      { text: "Điều 6. Phạt" },
      { text: "1. Mức phạt" },
      { text: "Phụ lục 01. Bảng giá" },
      { text: "Điều 5. Giá phụ lục" },
    ])
  );
  it("clauseKey chuẩn hóa nhiều cách gọi", () => {
    assert.strictEqual(clauseKey("Điều 5"), "5");
    assert.strictEqual(clauseKey("điều 5.2"), "5.2");
    assert.strictEqual(clauseKey("5.2(a)"), "5.2.a");
    assert.strictEqual(clauseKey("Điều 5.2 (a)"), "5.2.a");
    assert.strictEqual(clauseKey("khoản 2 Điều 5"), "5.2");
    assert.strictEqual(clauseKey("điểm a khoản 2 Điều 5"), "5.2.a");
    assert.strictEqual(clauseKey("khoản 5.2"), "5.2");
    assert.strictEqual(clauseKey("5.2.a"), "5.2.a");
  });
  it("findClause: đúng điều; nhiều điều trùng số thì trả hết; chương và phụ lục", () => {
    assert.deepStrictEqual(
      findClause(o, "Điều 5.2").map((x) => x.i),
      [3]
    );
    assert.deepStrictEqual(
      findClause(o, "điểm a khoản 2 điều 5").map((x) => x.i),
      [4]
    );
    assert.deepStrictEqual(
      findClause(o, "khoản 1 Điều 6").map((x) => x.i),
      [6]
    );
    assert.deepStrictEqual(
      findClause(o, "Điều 5").map((x) => x.i),
      [1, 8],
      "số trùng giữa thân hợp đồng và phụ lục"
    );
    assert.deepStrictEqual(
      findClause(o, "chương ii").map((x) => x.i),
      [0]
    );
    assert.deepStrictEqual(
      findClause(o, "Phụ lục 1").map((x) => x.i),
      [7]
    );
    assert.deepStrictEqual(findClause(o, "Điều 99"), []);
  });
  it("findClause: gọi 'Điều N' không khớp mục 'N.' của phụ lục (Chạy thật 23/11/2026, tasks/va-tra-dieu-khoan.md)", () => {
    const hd = buildOutline(
      rows([
        { text: "Điều 3. Thời hạn thuê" },
        { text: "3.1. Thời hạn thuê là 05 năm." },
        { text: "Điều 4. Tiền thuê" },
        { text: "PHỤ LỤC 01. MÔ TẢ DIỆN TÍCH" },
        { text: "1. Mô tả mặt bằng" },
        { text: "3. Giờ hoạt động" },
        { text: "3.1 Tòa Nhà mở cửa từ 6 giờ." },
      ])
    );
    assert.deepStrictEqual(findClause(hd, "Điều 3").map((x) => x.i), [0]);
    assert.deepStrictEqual(findClause(hd, "điều 3").map((x) => x.i), [0], "không phân biệt hoa thường");
    assert.deepStrictEqual(findClause(hd, "Article 3").map((x) => x.i), [0]);
    assert.deepStrictEqual(findClause(hd, "3").map((x) => x.i), [0, 5], "chỉ gọi số thì vẫn báo trùng để AI chọn");
    assert.deepStrictEqual(findClause(hd, "khoản 1 Điều 3").map((x) => x.i), [1, 6], "khoản trùng số giữ cách cũ: trả hết");
    // Phụ lục không có điều mang chữ "Điều": gọi "Điều 1" vẫn không lấy mục "1." của phụ lục
    assert.deepStrictEqual(findClause(hd, "Điều 1"), []);
    // Văn bản không có mục nào kiểu "Điều N" (đánh số 1., 2. ngay từ đầu): gọi "Điều 3" vẫn tìm ra mục "3."
    const tron = buildOutline(rows([{ text: "1. Phạm vi" }, { text: "2. Giá" }, { text: "3. Thanh toán" }]));
    assert.deepStrictEqual(findClause(tron, "Điều 3").map((x) => x.i), [2]);
  });
  it("clauseAt: điều khoản trong cùng chứa một đoạn", () => {
    assert.strictEqual(clauseAt(o, 4)?.num, "5.2(a)");
    assert.strictEqual(clauseAt(o, 3)?.num, "5.2");
    assert.strictEqual(clauseAt(o, 6)?.num, "6.1");
    assert.strictEqual(clauseAt(buildOutline(rows([{ text: "Lời mở đầu" }, { text: "Điều 1. A" }])), 0), null);
  });
});

describe("Đợt 1A – so khác mức từ (diffWords)", () => {
  const clause =
    "Bên B thanh toán cho Bên A trong thời hạn 30 ngày kể từ ngày nhận hóa đơn; nếu chậm quá 30 ngày thì Bên B chịu phạt 30 triệu đồng và lãi 0,05% mỗi ngày trên số tiền chậm trả, trừ trường hợp bất khả kháng theo Điều 12.";
  const check = (oldT: string, newT: string, n: number) => {
    const h = diffWords(oldT, newT);
    assert.strictEqual(applyHunks(oldT, h), newT, "áp lại phải ra đúng câu mới");
    assert.strictEqual(h.length, n, JSON.stringify(h));
    return h;
  };
  it("đổi một con số lặp lại: đúng một chỗ, xóa đúng 2 ký tự", () => {
    const h = check(clause, clause.replace("quá 30 ngày", "quá 45 ngày"), 1);
    assert.deepStrictEqual({ del: h[0].del, ins: h[0].ins }, { del: "30", ins: "45" });
    assert.strictEqual(h[0].at, clause.indexOf("quá 30") + 4);
  });
  it("hai chỗ sửa cách xa nhau: hai khối riêng", () => {
    const h = check(clause, clause.replace("thời hạn 30 ngày", "thời hạn 15 ngày").replace("0,05%", "0,03%"), 2);
    assert.deepStrictEqual(
      h.map((x) => `${x.del}>${x.ins}`),
      ["30>15", "0,05%>0,03%"]
    );
  });
  it("sửa đầu và cuối câu; thêm thuần; xóa thuần", () => {
    check(clause, clause.replace("Bên B thanh toán", "Bên Mua thanh toán").replace("Điều 12.", "Điều 14."), 2);
    const add = check(clause, clause.replace("thanh toán cho", "thanh toán đầy đủ cho"), 1);
    assert.strictEqual(add[0].del, "");
    assert.strictEqual(add[0].ins.trim(), "đầy đủ");
    const del = check(clause, clause.replace(", trừ trường hợp bất khả kháng theo Điều 12", ""), 1);
    assert.strictEqual(del[0].ins, "");
  });
  it("hai từ liền nhau đổi cùng lúc gộp thành một khối; câu giống hệt không có khối nào", () => {
    const h = check("chậm quá 30 ngày thì phạt", "chậm quá 45 hôm thì phạt", 1);
    assert.deepStrictEqual({ del: h[0].del, ins: h[0].ins }, { del: "30 ngày", ins: "45 hôm" });
    assert.deepStrictEqual(diffWords(clause, clause), []);
  });
  it("thử ngẫu nhiên 3.000 cặp câu: áp các khối khác biệt vào câu cũ luôn ra đúng câu mới; kế hoạch sửa thi hành ra đúng câu mới", () => {
    let seed = 12345;
    const rnd = (n: number) => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) % n;
    const words = ["Bên", "A", "B", "30", "ngày", "0,05%", "thanh", "toán", ".", ",", "và", "30", "ab", "a"];
    const sentence = () => Array.from({ length: 1 + rnd(9) }, () => words[rnd(words.length)]).join(rnd(5) ? " " : "  ");
    for (let t = 0; t < 3000; t++) {
      const oldT = sentence();
      const newT = rnd(4) ? sentence() : oldT.replace(words[rnd(words.length)], words[rnd(words.length)]);
      const h = diffWords(oldT, newT);
      assert.strictEqual(applyHunks(oldT, h), newT, `${JSON.stringify(oldT)} → ${JSON.stringify(newT)}`);
      const plan = planParagraphEdit(oldT, oldT, newT);
      if (typeof plan === "string") continue;
      // Thi hành kế hoạch trên chuỗi: mỗi phép định vị bằng lần xuất hiện thứ k
      const spots = plan.ops
        .map((op: any) => {
          if (op.kind === "insertStart") return { s: 0, e: 0, text: op.text };
          const s = op.from.at;
          const e = op.kind === "replace" ? (op.to ? op.to.at + op.to.search.length : s + op.from.search.length) : s + op.from.search.length;
          assert.strictEqual(oldT.substr(op.from.at, op.from.search.length), op.from.search);
          assert.strictEqual(occurrenceAt(oldT, op.from.search, op.from.at).index, op.from.occurrence);
          if (op.kind === "replace") return { s, e, text: op.text };
          return op.kind === "insertAfter" ? { s: e, e, text: op.text } : { s, e: s, text: op.text };
        })
        .sort((x: any, y: any) => y.s - x.s);
      let out = oldT;
      for (const sp of spots) out = out.slice(0, sp.s) + sp.text + out.slice(sp.e);
      assert.strictEqual(out, newT, `kế hoạch sai: ${JSON.stringify(oldT)} → ${JSON.stringify(newT)}`);
    }
  });
  it("số có dấu chấm, phẩy là một từ: 5.000.000.000 đổi thành 6.000.000.000 là một khối trọn số", () => {
    const h = check("Giá trị là 5.000.000.000 đồng.", "Giá trị là 6.000.000.000 đồng.", 1);
    assert.strictEqual(h[0].del, "5.000.000.000");
  });
});

describe("Đợt 1A – kế hoạch sửa đoạn trong Word (planParagraphEdit)", () => {
  it("đoạn sạch: đổi '30' thứ hai → tìm '30', lần thứ hai trong ba lần", () => {
    const C = "Thanh toán trong 30 ngày; chậm quá 30 ngày thì phạt 30 triệu.";
    const p = planParagraphEdit(C, C, C.replace("quá 30", "quá 45"));
    assert.ok(typeof p !== "string");
    if (typeof p === "string") return;
    assert.strictEqual(p.hunks, 1);
    assert.deepStrictEqual(p.ops, [{ kind: "replace", from: { search: "30", occurrence: 1, total: 3, at: C.indexOf("quá 30") + 4 }, text: "45" }]);
  });

  it("đoạn đang có chữ đối tác đã xóa: định vị trên chuỗi gồm cả chữ đã xóa, không đụng vào chữ đã xóa", () => {
    // Đúng dữ liệu thăm dò: .text = "…là 24 tháng12 tháng…", bản hiện tại = "…là 12 tháng…"
    const M = "Thời hạn thuê là 24 tháng12 tháng kể từ ngày ký.";
    const C = "Thời hạn thuê là 12 tháng kể từ ngày ký.";
    assert.strictEqual(typeof planParagraphEdit(M, C, "Thời hạn thuê là 18 tháng kể từ ngày bàn giao."), "string", "không có vị trí khúc đã xóa mà gán nhập nhằng thì từ chối");
    const p = planParagraphEdit(M, C, "Thời hạn thuê là 18 tháng kể từ ngày bàn giao.", [[17, 25]]);
    assert.ok(typeof p !== "string");
    if (typeof p === "string") return;
    assert.strictEqual(p.ops.length, 2);
    const [a, b] = p.ops as any[];
    assert.deepStrictEqual({ s: a.from.search, at: a.from.at, t: a.text }, { s: "12", at: 25, t: "18" });
    assert.strictEqual(b.from.search, "ký");
    assert.strictEqual(b.text, "bàn giao");
  });

  it("chỗ sửa vắt qua khúc đã xóa: tách thành từng khúc sống, chữ mới đặt vào khúc đầu", () => {
    const M = "Thời hạn thuê là 24 tháng12 tháng kể từ ngày ký.";
    const C = "Thời hạn thuê là 12 tháng kể từ ngày ký.";
    const p = planParagraphEdit(M, C, "Thời hạn thuê 36 tháng kể từ ngày ký.", [[17, 25]]);
    assert.ok(typeof p !== "string");
    if (typeof p === "string") return;
    const ops = p.ops as any[];
    assert.strictEqual(ops.length, 2);
    assert.deepStrictEqual(
      ops.map((o) => [o.kind, o.from.search, o.text]),
      [
        ["replace", "là ", "36"],
        ["replace", "12", ""],
      ]
    );
  });

  it("thêm thuần: neo vào từ đứng trước; chèn đầu đoạn; đoạn rỗng", () => {
    const C = "Bên B thanh toán trong 30 ngày.";
    const p1 = planParagraphEdit(C, C, "Bên B thanh toán đủ trong 30 ngày.") as any;
    assert.strictEqual(p1.ops[0].kind, "insertAfter");
    assert.strictEqual(p1.ops[0].from.search, "toán ");
    assert.strictEqual(p1.ops[0].text, "đủ ");
    const p2 = planParagraphEdit(C, C, "Theo đó, Bên B thanh toán trong 30 ngày.") as any;
    assert.strictEqual(p2.ops[0].kind, "insertBefore");
    assert.strictEqual(p2.ops[0].from.search, "Bên");
    const p3 = planParagraphEdit("", "", "Nội dung mới") as any;
    assert.deepStrictEqual(p3.ops, [{ kind: "insertStart", text: "Nội dung mới" }]);
  });

  it("khúc xóa dài hơn 255 ký tự: định vị bằng đầu và đuôi", () => {
    const long = "từ " + "rất dài ".repeat(60) + "hết";
    const C = `Mở đầu ${long} kết thúc.`;
    const p = planParagraphEdit(C, C, "Mở đầu kết thúc.") as any;
    assert.strictEqual(p.ops.length, 1);
    assert.ok(p.ops[0].to, "phải có điểm cuối");
    assert.ok(p.ops[0].from.search.length <= 255 && p.ops[0].to.search.length <= 255);
  });

  it("bản hiện tại không nằm trong chuỗi tìm kiếm: trả lỗi, không lập kế hoạch bừa", () => {
    assert.strictEqual(typeof planParagraphEdit("abc", "xyz", "xyz 1"), "string");
  });

  it("liveMap: dùng vị trí khúc đã xóa khi có; không có thì dò dãy con", () => {
    assert.deepStrictEqual(liveMap("ab ab", "ab", [[0, 3]]), [3, 4]);
    assert.strictEqual(liveMap("ab ab", "ab"), null, "không biết khúc nào đã xóa và có hai cách gán: từ chối");
    assert.strictEqual(liveMap("ab ab", "ab", [[9, 12]]), null, "vị trí sai thì dò dãy con, vẫn nhập nhằng");
    assert.deepStrictEqual(liveMap("xyz ab", "ab"), [4, 5], "chỉ một cách gán thì nhận");
    // Lỗi do lần soát mã độc lập tìm ra: chữ "1" của khúc đã xóa bị nhận là chữ sống
    const M2 = "là 12 tháng18 tháng";
    assert.strictEqual(typeof planParagraphEdit(M2, "là 18 tháng", "là 20 tháng"), "string");
    const ok = planParagraphEdit(M2, "là 18 tháng", "là 20 tháng", [[3, 11]]) as any;
    assert.deepStrictEqual({ s: ok.ops[0].from.search, at: ok.ops[0].from.at, t: ok.ops[0].text }, { s: "18", at: 11, t: "20" });
    assert.strictEqual(liveMap("abc", "abd"), null);
  });

  it("occurrenceAt đếm như Word tìm: trái sang phải, không chồng nhau", () => {
    assert.deepStrictEqual(occurrenceAt("30 và 30 và 30", "30", 6), { index: 1, total: 3 });
    assert.deepStrictEqual(occurrenceAt("aaa", "aa", 1), { index: -1, total: 1 });
  });
});

describe("Đợt 1A – chống sửa nhầm đoạn (locateParagraph)", () => {
  const doc = ["Tiêu đề", "Điều 1. Thời hạn thuê là 12 tháng.", "Điều 2. Giá thuê là 20 triệu đồng.", "Đại diện: ………", "Đại diện: ………"];
  it("đúng chỉ số và khớp mấy chữ đầu: dùng luôn", () => {
    assert.deepStrictEqual(locateParagraph(doc, 2, "Điều 2. Giá thuê"), { ok: true, index: 2, moved: false });
  });
  it("chèn hai đoạn phía trên rồi sửa theo số đoạn cũ: tự định vị lại đúng đoạn", () => {
    const shifted = ["Tiêu đề", "Đoạn mới 1", "Đoạn mới 2", ...doc.slice(1)];
    assert.deepStrictEqual(locateParagraph(shifted, 2, "Điều 2. Giá thuê"), { ok: true, index: 4, moved: true });
  });
  it("không còn đoạn nào khớp: từ chối", () => {
    const r = locateParagraph(doc, 2, "Điều 9. Không có");
    assert.strictEqual(r.ok, false);
    assert.ok(!r.ok && /Đọc lại/.test(r.reason));
  });
  it("nhiều đoạn giống nhau mà chỉ số lệch: từ chối kèm ứng viên, không tự chọn", () => {
    const r = locateParagraph(doc, 1, "Đại diện:");
    assert.strictEqual(r.ok, false);
    assert.ok(!r.ok && r.candidates.join(",") === "3,4");
    // đúng chỉ số thì vẫn dùng được dù có đoạn trùng
    assert.deepStrictEqual(locateParagraph(doc, 4, "Đại diện:"), { ok: true, index: 4, moved: false });
  });
  it("không có expect thì theo chỉ số; chỉ số ngoài phạm vi thì từ chối", () => {
    assert.deepStrictEqual(locateParagraph(doc, 1), { ok: true, index: 1, moved: false });
    assert.strictEqual(locateParagraph(doc, 99).ok, false);
  });
  it("matchesExpect bỏ qua khác biệt khoảng trắng và dấu … cuối", () => {
    assert.ok(matchesExpect("Điều 1.  Thời hạn thuê", "Điều 1. Thời hạn…"));
    assert.ok(!matchesExpect("Điều 1. Thời hạn", "Điều 2."));
  });
});

describe("Đợt 1A – Hoàn tác cho Word: sửa đổi nào là của lượt vừa rồi", () => {
  const k = (type: string, author: string, text: string, date = "2026-11-15T12:03:00Z") => changeKey({ type, author, date, text });
  it("phần chênh giữa hai lần chụp là sửa đổi của AI; sửa đổi của đối tác không bị tính", () => {
    const before = [k("Deleted", "Đối tác", "24 tháng"), k("Added", "Đối tác", "12 tháng")];
    const after = [k("Deleted", "Đối tác", "24 tháng"), k("Added", "Tôi", "18"), k("Deleted", "Tôi", "12"), k("Added", "Đối tác", "12 tháng")];
    assert.deepStrictEqual(addedKeys(before, after), [k("Added", "Tôi", "18"), k("Deleted", "Tôi", "12")]);
    assert.deepStrictEqual(addedKeys(after, after), []);
  });
  it("sửa đổi trùng hệt (cùng người, cùng phút, cùng chữ): chỉ phần thêm mới được tính, đúng vị trí", () => {
    const a = k("Added", "Tôi", "45");
    const b = k("Deleted", "Tôi", "30");
    assert.deepStrictEqual(addedKeys([a, b], [a, b, a, b]), [a, b]);
    assert.deepStrictEqual(pickByKeys([a, b, a, b], [a, b]), [3, 2], "từ chối cặp đứng sau");
  });
  it("Word gộp phần chèn mới vào phần chèn liền kề có từ trước: không được từ chối (sẽ mất cả phần cũ), phải hoàn tác cách khác", () => {
    const s = (type: string, author: string, text: string) => snapChange({ type, author, date: "2026-11-15T12:03:00Z", text });
    const old = s("Added", "Tôi", "\rKhoản chèn lượt trước.");
    const partner = s("Deleted", "Đối tác", "24 tháng");
    // Lượt này chèn một khoản sát ngay trước: Word gộp thành một sửa đổi; đồng thời sửa một chữ ở đoạn khác
    const mergedNow = s("Added", "Tôi", "\rKhoản chèn lượt này.\rKhoản chèn lượt trước.");
    const d = diffChanges([partner, old], [partner, s("Added", "Tôi", "45"), s("Deleted", "Tôi", "30"), mergedNow]);
    assert.deepStrictEqual(d.added, [s("Added", "Tôi", "45").key, s("Deleted", "Tôi", "30").key]);
    assert.deepStrictEqual(d.merged, [mergedNow.key]);
    assert.strictEqual(d.removed, 1);
    // Sửa bên trong phần chèn đang chờ duyệt: sửa đổi cũ đổi nội dung, không có sửa đổi mới hẳn
    const edited = s("Added", "Tôi", "\rKhoản chèn lượt trước, đã sửa.");
    const d2 = diffChanges([partner, old], [partner, edited]);
    assert.deepStrictEqual({ added: d2.added, merged: d2.merged, removed: d2.removed }, { added: [], merged: [edited.key], removed: 1 });
    // Không có gì bị gộp: mọi sửa đổi mới đều từ chối được
    const d3 = diffChanges([partner], [partner, s("Added", "Tôi", "45")]);
    assert.deepStrictEqual({ merged: d3.merged, removed: d3.removed, n: d3.added.length }, { merged: [], removed: 0, n: 1 });
  });
  it("sửa đổi của AI trùng khóa với sửa đổi người dùng tự làm ở đoạn khác: hoàn tác đúng cái nằm trong đoạn AI đã đụng", () => {
    const a = k("Deleted", "Tôi", "30");
    // Văn bản: [đoạn P1: của AI] rồi [đoạn P3: của người dùng] — cùng khóa
    const cur = [
      { key: a, para: "P1" },
      { key: a, para: "P3" },
    ];
    assert.deepStrictEqual(pickWanted(cur, [{ key: a, para: "P1" }]), [0], "chọn theo đoạn, không lấy mục đứng sau");
    assert.deepStrictEqual(pickWanted(cur, [{ key: a }]), [1], "không biết đoạn thì lấy mục đứng sau");
    assert.deepStrictEqual(pickWanted(cur, [{ key: a, para: "P9" }]), [1], "đoạn không còn thì chọn theo khóa");
    assert.deepStrictEqual(pickWanted(cur, [{ key: a, para: "P1" }, { key: a }]), [1, 0]);
    assert.deepStrictEqual(pickWanted([{ key: "x" }], [{ key: a }]), []);
  });
  it("duyệt tất cả sửa đổi của một tác giả: phải khớp đúng một tên", () => {
    const names = ["Nguyễn Văn An", "Hoàng Anh", "An Phát Law", "Đối tác"];
    assert.deepStrictEqual(resolveAuthor(names, "An"), ["Nguyễn Văn An", "Hoàng Anh", "An Phát Law"], "khớp nhiều tên: không chọn");
    assert.strictEqual(resolveAuthor(names, "hoàng anh"), "Hoàng Anh");
    assert.strictEqual(resolveAuthor(names, "đối tác"), "Đối tác");
    assert.strictEqual(resolveAuthor(names, "Phát"), "An Phát Law", "chỉ một tên chứa chuỗi đó");
    assert.deepStrictEqual(resolveAuthor(names, "Bình"), []);
    assert.strictEqual(resolveAuthor(["An", "Hoàng An"], "an"), "An", "trùng hẳn tên được ưu tiên");
  });
  it("pickByKeys: khóa không còn trong văn bản thì bỏ qua", () => {
    const a = k("Added", "Tôi", "45");
    assert.deepStrictEqual(pickByKeys([k("Added", "Đối tác", "x")], [a]), []);
  });
});

describe("Đợt 1A – công cụ và lời nhắc", () => {
  it("các công cụ mới có mặt; công cụ sửa và xóa theo đoạn bắt buộc expect", () => {
    const names = WORD_SPECS.map((s) => s.name);
    for (const n of ["word_get_outline", "word_find", "word_read_clause", "word_insert_clause", "word_change_level"]) assert.ok(names.includes(n), n);
    for (const n of ["word_set_paragraph_text", "word_delete_paragraphs", "word_insert_clause", "word_change_level"]) {
      const sp = WORD_SPECS.find((s) => s.name === n);
      assert.ok(sp?.parameters.required.includes("expect"), `${n} phải bắt buộc expect`);
    }
    const rep = WORD_SPECS.find((s) => s.name === "word_replace_text");
    for (const p of ["paragraph_index", "occurrence", "all"]) assert.ok(rep?.parameters.properties[p], `word_replace_text thiếu ${p}`);
  });
  it("lời nhắc hệ thống Word cố định (không ngày giờ, không trạng thái) và có hướng dẫn mục lục, expect", () => {
    assert.ok(/word_get_outline/.test(WORD_SYSTEM_PROMPT) && /expect/.test(WORD_SYSTEM_PROMPT));
    assert.ok(!/\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{4}/.test(WORD_SYSTEM_PROMPT));
  });
  it("mọi tên công cụ word_* nêu trong kỹ năng dựng sẵn đều tồn tại", () => {
    const names = new Set(WORD_SPECS.map((s) => s.name));
    for (const sk of BUILTIN_SKILLS) {
      const all = [sk.body, ...Object.values(sk.files || {})].join("\n");
      for (const m of all.match(/word_[a-z_]+/g) || []) assert.ok(names.has(m), `Kỹ năng ${sk.name} nhắc công cụ không tồn tại: ${m}`);
    }
  });
  it("'giai đoạn 2', 'công đoạn 4' không thành liên kết tới đoạn; [¶12] và 'đoạn 5' vẫn là liên kết", () => {
    const refs = (t: string) =>
      findRefs(t, "Word")
        .filter((p) => p.ref)
        .map((p) => (p.ref as any).index);
    assert.deepStrictEqual(refs("Dự án giai đoạn 2, công đoạn 4 và phân đoạn 7 đang triển khai"), []);
    assert.deepStrictEqual(refs("Xem [¶12] và đoạn 5, đoạn số 7"), [12, 5, 7]);
  });
  it("tìm trong văn bản: bỏ qua đoạn đã xóa, đếm số lần, nguyên từ, không dấu", () => {
    const rs = [
      { i: 0, text: "Bên B thanh toán trong 30 ngày, chậm thanh toán thì chịu lãi." },
      { i: 1, text: "", deleted: true },
      { i: 2, text: "Điều khoản về Thanh Toán bổ sung." },
      { i: 3, text: "Bên Bán giao hàng." },
    ];
    const r = searchRows(rs, "thanh toán");
    assert.deepStrictEqual(r, { total: 3, hits: [{ i: 0, at: 6, len: 10, n: 2 }, { i: 2, at: 14, len: 10, n: 1 }] });
    assert.strictEqual(searchRows(rs, "thanh toán", { matchCase: true }).total, 2);
    assert.strictEqual(searchRows(rs, "Bên B", { wholeWord: true }).total, 1, "không khớp 'Bên Bán'");
    assert.strictEqual(searchRows(rs, "Bên B").total, 2);
    assert.strictEqual(searchRows(rs, "thanh toan").total, 0);
    assert.strictEqual(searchRows(rs, "thanh toan", { ignoreAccents: true }).total, 3);
  });
  it("không đoạn nào chứa nguyên cụm: tìm theo từng từ, đoạn có nhiều từ của câu hỏi nhất đứng trước", () => {
    const rs = [
      { i: 0, text: "Tiền Thuê được thanh toán theo quý." },
      { i: 1, text: "Nếu chậm thanh toán, Bên Thuê phải trả lãi chậm trả 0,05% mỗi ngày. Câu khác không liên quan." },
      { i: 2, text: "Bên Cho Thuê xuất hóa đơn." },
      { i: 3, text: "Lãi không tính trên tiền đặt cọc.", deleted: false },
    ];
    assert.strictEqual(searchRows(rs, "lãi chậm thanh toán").total, 0);
    const w = searchRowsByWords(rs, "lãi chậm thanh toán");
    assert.deepStrictEqual(w.words, ["lai", "cham", "thanh", "toan"]);
    assert.deepStrictEqual(
      w.hits.map((h) => [h.i, h.matched]),
      [
        [1, 4],
        [0, 2],
      ]
    );
    assert.ok(w.hits[0].sentence.startsWith("Nếu chậm thanh toán"), "trích câu chứa nhiều từ khớp nhất");
    assert.deepStrictEqual(searchRowsByWords(rs, "lãi").hits, [], "một từ thì không tìm theo từng từ");
  });
  it("tìm và thay một cụm: chỉ đổi đúng từ khác đi khi cụm cũ và mới có phần chung", () => {
    const p = planParagraphEdit("chưa gồm thuế giá trị gia tăng", "chưa gồm thuế giá trị gia tăng", "đã gồm thuế giá trị gia tăng") as any;
    assert.strictEqual(p.hunks, 1);
    assert.deepStrictEqual({ s: p.ops[0].from.search, t: p.ops[0].text }, { s: "chưa", t: "đã" });
    assert.ok(p.ratio < 0.8);
    const q = planParagraphEdit("30", "30", "60") as any;
    assert.ok(q.ratio >= 0.8, "cụm khác hẳn thì thay cả cụm");
  });
  it("dòng '5. Thanh toán' đứng lẻ giữ nguyên số; hai dòng 1. 2. mới thành danh sách", () => {
    assert.strictEqual(isMarkdownNumberedList(["5. Thanh toán"]), false);
    assert.strictEqual(isMarkdownNumberedList(["2026. Năm tài chính", "Nội dung"]), false);
    assert.strictEqual(isMarkdownNumberedList(["1. Một", "2. Hai"]), true);
    assert.deepStrictEqual(parseMarkdownLine("5. Thanh toán", false), { text: "5. Thanh toán", style: "Normal" });
    assert.deepStrictEqual(parseMarkdownLine("1. Một", true), { text: "Một", style: "ListNumber" });
    assert.deepStrictEqual(parseMarkdownLine("## **Điều 5**"), { text: "Điều 5", style: "Heading2" });
  });
  it("kết quả đọc đoạn kèm số điều, cờ đang có sửa đổi, cờ đã xóa", () => {
    const r: any = packParagraphs(
      [
        { text: "Thanh toán", style: "Heading 1", num: "Điều 5" },
        { text: "Bên B trả tiền", style: "Normal", changed: true },
        { text: "", style: "Normal", deleted: true },
      ],
      0,
      3
    );
    assert.deepStrictEqual(r.paragraphs[0], { i: 0, num: "Điều 5", style: "Heading 1", text: "Thanh toán" });
    assert.strictEqual(r.paragraphs[1].changed, true);
    assert.strictEqual(r.paragraphs[2].deleted, true);
    assert.strictEqual(r.paragraphs[1].num, undefined);
  });
});
