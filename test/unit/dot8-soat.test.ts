/** Đợt 8 — các chỗ sửa theo lượt soát mã độc lập (tasks/dot/dot-8/bang-chung/soat-ma-dot-8.md) */
import * as assert from "assert";
import { controlRefusal, fillControlOoxml, findBlanks, formatDate } from "../../src/core/formfill";
import { checkConsistency } from "../../src/core/consistency";
import { buildOutline } from "../../src/core/outline";

const wrap = (sdt: string) => `<pkg:package><w:p><w:r><w:t>Trước </w:t></w:r>${sdt}</w:p></pkg:package>`;
/** Đếm thẻ mở / đóng một tên thẻ: OOXML dựng ra phải cân */
const balanced = (x: string, tag: string) => (x.match(new RegExp(`<${tag}(?=[\\s>])`, "g")) || []).length === (x.match(new RegExp(`</${tag}>`, "g")) || []).length;

describe("Đợt 8 · soát mã — chỗ trống", () => {
  it("#8 Nhận thêm: dãy chấm dính chữ, năm 20……, …/…/20…, ngày.....tháng.....năm", () => {
    const has = (t: string) => findBlanks(t).length > 0;
    assert.ok(has("Họ và tên.........."), "Họ và tên..........");
    assert.ok(has("Hà Nội, ngày 5 tháng 3 năm 20……"), "năm 20……");
    assert.ok(has("Ngày sinh: …/…/20…"), "…/…/20…");
    assert.ok(has("ngày.....tháng.....năm 20....."), "ngày.....tháng.....");
    assert.ok(has("Sinh năm 19…"), "19…");
  });
  it("#9 Không nhận nhầm: dẫn mục lục, (…), “…” trích dẫn, cả dòng gạch, dấu ba chấm cuối câu", () => {
    const none = (t: string) => assert.deepStrictEqual(findBlanks(t), [], t);
    none("Điều 1. Giải thích từ ngữ ........................ 5");
    none("Chương II ……………… 12");
    none("Các loại giấy tờ (…) theo quy định.");
    none("Bên A cam kết “…” như đã thỏa thuận.");
    none("____________________________");
    none("Các bên thống nhất như sau...");
    none("Năm 2026 doanh thu tăng.");
  });
  it("#9 Soát nhất quán dùng cùng bộ loại trừ: dẫn mục lục không bị báo là chỗ trống", () => {
    const rows = [{ i: 0, text: "Điều 1. Giải thích từ ngữ ........................ 5" }, { i: 1, text: "Họ và tên: ………………" }];
    const r = checkConsistency(rows as any, { checks: ["blank"], outline: buildOutline(rows as any) });
    assert.deepStrictEqual(
      r.issues.map((x) => x.p),
      [1]
    );
  });
  it("#10 Dấu cách không ngắt: chuỗi tìm giữ nguyên dấu cách không ngắt của văn bản", () => {
    const t = "Hà Nội, ngày   tháng   năm 2026";
    const b = findBlanks(t);
    assert.strictEqual(b.length, 1);
    assert.ok(t.includes(b[0].str), "chuỗi tìm phải có nguyên trong chữ của đoạn");
    assert.ok(b[0].str.includes(" "));
  });
});

describe("Đợt 8 · soát mã — dựng OOXML của ô", () => {
  it("#5 Run có sửa định dạng chờ duyệt (w:rPrChange): OOXML dựng ra cân thẻ", () => {
    const sdt =
      '<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="X"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>cũ</w:t></w:r></w:sdtContent></w:sdt>';
    const x = fillControlOoxml(wrap(sdt), { text: "mới" })!;
    for (const tag of ["w:r", "w:rPr", "w:sdt", "w:sdtContent", "w:p"]) assert.ok(balanced(x, tag), `${tag} không cân: ${x}`);
    assert.match(x, /<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">mới<\/w:t><\/w:r>/);
    const blk =
      '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="4" w:author="X"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>a</w:t></w:r></w:p></w:sdtContent></w:sdt>';
    const y = fillControlOoxml(`<pkg:package>${blk}</pkg:package>`, { text: "b" })!;
    for (const tag of ["w:p", "w:pPr", "w:r"]) assert.ok(balanced(y, tag), `${tag} không cân: ${y}`);
  });
  it("#6 #14 #17 Ô không nên ghi đè: chứa bảng / hình, gắn dữ liệu, bao ô bảng — có lý do, không dựng", () => {
    const tbl = '<w:sdt><w:sdtPr/><w:sdtContent><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>';
    const pic = '<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:drawing>…</w:drawing></w:r></w:sdtContent></w:sdt>';
    const bound = '<w:sdt><w:sdtPr><w:dataBinding w:xpath="/a"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>';
    const cell = '<w:sdt><w:sdtPr/><w:sdtContent><w:tc><w:p/></w:tc></w:sdtContent></w:sdt>';
    for (const [sdt, re] of [
      [tbl, /bảng/],
      [pic, /hình/],
      [bound, /dữ liệu/],
      [cell, /ô bảng/],
    ] as [string, RegExp][]) {
      assert.match(controlRefusal(wrap(sdt)) || "", re);
      assert.strictEqual(fillControlOoxml(wrap(sdt), { text: "x" }), null);
    }
    assert.strictEqual(controlRefusal(wrap('<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>a</w:t></w:r></w:sdtContent></w:sdt>')), null);
  });
  it("#13 Ô đang hiện chữ gợi ý: bỏ mọi kiểu chữ của run (tên kiểu gợi ý khác nhau theo ngôn ngữ)", () => {
    const sdt = '<w:sdt><w:sdtPr><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="VnbnGichCh"/></w:rPr><w:t>Nhập</w:t></w:r></w:sdtContent></w:sdt>';
    assert.ok(!/rStyle/.test(fillControlOoxml(wrap(sdt), { text: "x" })!));
  });
  it("#16 Ký tự điều khiển bị thay, tab thành w:tab", () => {
    const sdt = '<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>a</w:t></w:r></w:sdtContent></w:sdt>';
    const x = fillControlOoxml(wrap(sdt), { text: "A\u000Bb\u0007c\tD" })!;
    assert.ok(!/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(x));
    assert.match(x, /<w:tab\/><w:t xml:space="preserve">D<\/w:t>/);
  });
  it("#17 Ô chứa đoạn dạng <w:p/>: nhận là ô chứa đoạn", () => {
    const x = fillControlOoxml('<pkg:package><w:sdt><w:sdtPr/><w:sdtContent><w:p/></w:sdtContent></w:sdt></pkg:package>', { text: "Một" })!;
    assert.match(x, /<w:sdtContent><w:p><w:r><w:t xml:space="preserve">Một<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/);
  });
  it("#15 Định dạng ngày có chữ trong ngoặc hoặc giờ phút: viết dd/MM/yyyy", () => {
    assert.strictEqual(formatDate({ d: 5, m: 3, y: 2027 }, "'ngày' dd 'tháng' MM"), "05/03/2027");
    assert.strictEqual(formatDate({ d: 5, m: 3, y: 2027 }, "dd/MM/yyyy HH:mm"), "05/03/2027");
    assert.strictEqual(formatDate({ d: 5, m: 3, y: 2027 }, "d.M.yyyy"), "5.3.2027");
  });
});

import { cellLabel, emptyCells, parseFieldId as pid } from "../../src/office/wordForm";
import { WORD_SPECS as SPECS } from "../../src/office/wordTools";

describe("Đợt 8 · soát mã — ô bảng, nhãn bắt buộc", () => {
  it("#20 Bảng nhiều cột: ô trống nhận nhãn từ ô có chữ gần nhất bên trái + tên cột; #21 không lấy cột đầu", () => {
    const v = [
      ["Hạng mục", "Diện tích", "Đơn giá", "Thành tiền"],
      ["Căn A", "", "", ""],
      ["", "", "", ""],
    ];
    assert.deepStrictEqual(
      emptyCells(v).map((x) => `${x.r}.${x.c} ${x.label}`),
      ["1.1 Căn A — Diện tích", "1.2 Căn A — Đơn giá", "1.3 Căn A — Thành tiền"]
    );
    assert.strictEqual(cellLabel(v, 2, 1), null, "hàng trống hết không có nhãn");
  });
  it("#1 label là tham số bắt buộc của word_fill_form", () => {
    const sp = SPECS.find((x) => x.name === "word_fill_form")!;
    assert.deepStrictEqual((sp.parameters as any).properties.fields.items.required, ["id", "value", "label"]);
    assert.ok(pid("b3.0"));
  });
});
