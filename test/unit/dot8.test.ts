/** Đợt 8 — điền mẫu trong Word: phần thuần (nhận diện chỗ trống, nhãn, lần xuất hiện; OOXML ô chọn / ô ngày; giá trị điền) */
import * as assert from "assert";
import {
  checkboxValue,
  countOccurrences,
  dateFormatFromOoxml,
  findBlanks,
  formatDate,
  listItemsFromOoxml,
  matchListItem,
  occurrenceOf,
  parseDate,
  selectDropdownOoxml,
} from "../../src/core/formfill";

describe("Đợt 8 · chỗ trống gõ tay", () => {
  it("Dãy chấm dài: một chỗ trống, đúng trọn dãy, nhãn là chữ đứng trước", () => {
    const b = findBlanks("Họ và tên người đại diện Bên B: ……………………………………");
    assert.strictEqual(b.length, 1);
    assert.strictEqual(b[0].str, "……………………………………");
    assert.strictEqual(b[0].label, "Họ và tên người đại diện Bên B:");
    assert.strictEqual(b[0].occurrence, 0);
  });
  it("Nhiều chỗ trống trong một đoạn: nhãn của chỗ sau là chữ giữa hai chỗ", () => {
    const b = findBlanks("Số CCCD: ____________________ cấp ngày ………… tại …………………");
    assert.deepStrictEqual(
      b.map((x) => x.label),
      ["Số CCCD:", "cấp ngày", "tại"]
    );
  });
  it("'[●]' lặp hai lần: lần xuất hiện 0 và 1", () => {
    const b = findBlanks("Số tiền đặt cọc: [●] đồng (bằng chữ: [●] đồng).");
    assert.deepStrictEqual(
      b.map((x) => [x.str, x.occurrence, x.label]),
      [
        ["[●]", 0, "Số tiền đặt cọc:"],
        ["[●]", 1, "đồng (bằng chữ:"],
      ]
    );
  });
  it("'ngày   tháng   năm' là một chỗ trống", () => {
    const b = findBlanks("Hà Nội, ngày   tháng   năm 2026");
    assert.strictEqual(b.length, 1);
    assert.match(b[0].str, /^ngày\s+tháng\s+năm$/);
  });
  it("Không nhận nhầm: dòng ký tên, dấu ba chấm sau chữ, số thập phân, câu thường", () => {
    assert.deepStrictEqual(findBlanks("(Ký, ghi rõ họ tên)"), []);
    assert.deepStrictEqual(findBlanks("ĐẠI DIỆN BÊN A ……………… ĐẠI DIỆN BÊN B (ký, đóng dấu)"), []);
    assert.deepStrictEqual(findBlanks("Các bên thống nhất như sau..."), []);
    assert.deepStrictEqual(findBlanks("Diện tích 85.5 m², giá 1.250.000.000 đồng."), []);
    assert.deepStrictEqual(findBlanks("Bên A có trách nhiệm bàn giao nhà đúng hạn."), []);
  });
  it("Lần xuất hiện đếm không chồng nhau như Word tìm", () => {
    assert.strictEqual(occurrenceOf("…… và …… và ……", "……", 6), 1);
    assert.strictEqual(occurrenceOf("…… và …… và ……", "……", 8), 2);
    // Chỗ trống ngắn đứng sau một dãy dài: Word tìm "……" cũng trúng các đoạn nằm trong dãy dài → lần xuất hiện tính cả các đoạn đó
    const b = findBlanks("Tên: ………… tuổi: ……");
    assert.deepStrictEqual(
      b.map((x) => [x.str, x.occurrence]),
      [
        ["…………", 0],
        ["……", 2],
      ]
    );
    assert.strictEqual(countOccurrences("[●] x [●] y [●]", "[●]"), 3);
    // Chuỗi ngắn hơn nằm trong dãy dài: Word cũng đếm như vậy
    assert.strictEqual(countOccurrences("…………", "……"), 2);
  });
});

const OOXML_DROP =
  '<pkg:package><w:document><w:body><w:p><w:sdt><w:sdtPr><w:alias w:val="Hình thức thanh toán"/><w:tag w:val="thanh_toan"/><w:id w:val="8005"/><w:showingPlcHdr/>' +
  '<w:dropDownList><w:listItem w:displayText="Chuyển khoản" w:value="Chuyển khoản"/><w:listItem w:displayText="Tiền mặt" w:value="Tiền mặt"/><w:listItem w:displayText="A &amp; B" w:value="AB"/></w:dropDownList></w:sdtPr>' +
  '<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Chọn một mục</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document></pkg:package>';

describe("Đợt 8 · ô điền mẫu đọc qua OOXML", () => {
  it("Danh sách mục của ô chọn (giải mã &amp;)", () => {
    assert.deepStrictEqual(listItemsFromOoxml(OOXML_DROP), ["Chuyển khoản", "Tiền mặt", "A & B"]);
  });
  it("Chọn mục: bỏ chữ gợi ý và kiểu PlaceholderText, giữ phần ngoài ô", () => {
    const x = selectDropdownOoxml(OOXML_DROP, "Tiền mặt")!;
    assert.ok(x.startsWith("<pkg:package>") && x.endsWith("</pkg:package>"));
    assert.ok(!/showingPlcHdr/.test(x));
    assert.ok(!/PlaceholderText/.test(x));
    assert.match(x, /<w:sdtContent><w:r><w:rPr><\/w:rPr><w:t xml:space="preserve">Tiền mặt<\/w:t><\/w:r><\/w:sdtContent>/);
    assert.strictEqual(listItemsFromOoxml(x).length, 3, "danh sách mục còn nguyên");
  });
  it("Định dạng ngày của ô ngày", () => {
    assert.strictEqual(dateFormatFromOoxml('<w:sdt><w:sdtPr><w:date><w:dateFormat w:val="dd/MM/yyyy"/></w:date></w:sdtPr></w:sdt>'), "dd/MM/yyyy");
    assert.strictEqual(dateFormatFromOoxml("<w:sdt><w:sdtPr><w:text/></w:sdtPr></w:sdt>"), null);
  });
});

describe("Đợt 8 · giá trị điền", () => {
  it("Ô đánh dấu", () => {
    assert.strictEqual(checkboxValue(true), true);
    assert.strictEqual(checkboxValue("có"), true);
    assert.strictEqual(checkboxValue("x"), true);
    assert.strictEqual(checkboxValue("không"), false);
    assert.strictEqual(checkboxValue("có thể"), null);
  });
  it("Ô chọn chỉ nhận mục có trong danh sách (không phân biệt hoa thường)", () => {
    assert.strictEqual(matchListItem(["Chuyển khoản", "Tiền mặt"], "chuyển  khoản"), "Chuyển khoản");
    assert.strictEqual(matchListItem(["Chuyển khoản", "Tiền mặt"], "Trả góp"), null);
  });
  it("Ngày: đọc nhiều kiểu, từ chối ngày không có thật, viết theo định dạng của ô", () => {
    assert.deepStrictEqual(parseDate("18/11/2026"), { d: 18, m: 11, y: 2026 });
    assert.deepStrictEqual(parseDate("2026-11-18"), { d: 18, m: 11, y: 2026 });
    assert.deepStrictEqual(parseDate("ngày 5 tháng 3 năm 2027"), { d: 5, m: 3, y: 2027 });
    assert.strictEqual(parseDate("31/02/2026"), null);
    assert.strictEqual(parseDate("11/18/2026"), null, "kiểu Mỹ tháng 18 không có");
    assert.strictEqual(formatDate({ d: 5, m: 3, y: 2027 }, "dd/MM/yyyy"), "05/03/2027");
    assert.strictEqual(formatDate({ d: 5, m: 3, y: 2027 }, "d/M/yy"), "5/3/27");
    assert.strictEqual(formatDate({ d: 5, m: 3, y: 2027 }, "dddd, MMMM d, yyyy"), "05/03/2027");
  });
});

import { BUILTIN_SKILLS, autoSelectSkills } from "../../src/core/skillpack";
import { WORD_SPECS, WORD_SYSTEM_PROMPT } from "../../src/office/wordTools";
import { parseFieldId } from "../../src/office/wordForm";

describe("Đợt 8 · công cụ và kỹ năng", () => {
  const word = BUILTIN_SKILLS.filter((s) => !s.hosts || s.hosts.includes("Word"));
  it("Kỹ năng điền mẫu tự nạp đúng câu; câu khác không kéo nó vào", () => {
    assert.strictEqual(autoSelectSkills("Điền giúp tôi hợp đồng đặt cọc này với thông tin sau", word)[0]?.name, "dien-mau");
    assert.strictEqual(autoSelectSkills("điền vào mẫu tờ khai theo giấy tờ đính kèm", word)[0]?.name, "dien-mau");
    assert.ok(!autoSelectSkills("Rà soát hợp đồng mua bán này giúp tôi", word).some((s) => s.name === "dien-mau"));
    assert.ok(!autoSelectSkills("Soạn công văn gửi Sở Xây dựng", word).some((s) => s.name === "dien-mau"));
  });
  it("Hai công cụ: liệt kê chỉ đọc, điền là công cụ ghi; lời nhắc hệ thống nhắc cả hai; lược đồ không có enum số (Gemini)", () => {
    const list = WORD_SPECS.find((x) => x.name === "word_list_form_fields")!;
    const fill = WORD_SPECS.find((x) => x.name === "word_fill_form")!;
    assert.ok(list && !list.write);
    assert.ok(fill && fill.write);
    assert.ok(WORD_SYSTEM_PROMPT.includes("word_list_form_fields") && WORD_SYSTEM_PROMPT.includes("word_fill_form"));
    assert.ok(!/"enum":\[\d/.test(JSON.stringify([list, fill])));
  });
  it("Mã chỗ điền", () => {
    assert.deepStrictEqual(parseFieldId("c8001"), { kind: "control", cc: 8001 });
    assert.deepStrictEqual(parseFieldId("b12.1"), { kind: "blank", i: 12, k: 1 });
    assert.deepStrictEqual(parseFieldId("t0.2.1"), { kind: "cell", t: 0, r: 2, c: 1 });
    assert.strictEqual(parseFieldId("12"), null);
    assert.strictEqual(parseFieldId("b12"), null);
  });
});
