/** Chặn bằng máy các lỗi đã gặp từ hai lần trở lên (nguyên tắc của anh Dũng 25/11/2026: lỗi lặp lại thì chặn hẳn). */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import { blockedReason } from "../../src/core/memory";

const SRC = path.resolve(__dirname, "../../src");
const walk = (d: string): string[] =>
  fs.readdirSync(d, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walk(path.join(d, e.name)) : /\.tsx?$/.test(e.name) ? [path.join(d, e.name)] : []));

describe("Chặn lỗi lặp lại", () => {
  // Gặp ở adminDoc.ts, textcheck.ts (4.11.11), pptStyle.ts và memory.ts (27/11/2026): \b của JavaScript chỉ hiểu chữ ASCII,
  // nên \b sát chữ tiếng Việt có dấu (đ, ơ, ư, á…) không khớp như mong đợi. Dùng (?<!\p{L}) / (?!\p{L}) kèm cờ u.
  it("không có \\b sát chữ tiếng Việt có dấu trong biểu thức chính quy của src/", () => {
    const VI = "àáảãạăắằẳẵặâấầẩẫậđèéẻẽẹêếềểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵ";
    const V = `[${VI}]`;
    // \b sát chữ có dấu, hoặc sát nhóm ( … ) / [ … ] có chữ có dấu: \b(đồng|…), (tỷ|triệu)\b, [đĐ]\b (phản biện mức 3)
    const bad = new RegExp(
      [`\\\\b${V}`, `${V}\\\\b`, `\\\\b\\((?:\\?:)?[^()]*${V}`, `${V}[^()]*\\)\\\\b`, `\\\\b\\[[^\\]]*${V}`, `${V}[^\\]]*\\]\\\\b`].join("|"),
      "i"
    );
    const hits: string[] = [];
    for (const f of walk(SRC)) {
      fs.readFileSync(f, "utf8")
        .split(/\r?\n/)
        .forEach((line, i) => {
          // Bỏ dòng chú thích nguyên dòng (// …, /* …, " * …"); không cắt giữa dòng (dễ cắt nhầm chuỗi chứa //)
          const code = /^\s*(\/\/|\/\*|\*\s|\*$)/.test(line) ? "" : line;
          if (bad.test(code)) hits.push(`${path.relative(SRC, f)}:${i + 1}`);
        });
    }
    assert.deepStrictEqual(hits, [], "Thay \\b bằng (?<!\\p{L}) / (?!\\p{L}) và thêm cờ u: " + hits.join(", "));
  });

  it("bộ nhớ: dãy số theo sau là 'đ' được coi là số tiền, không chặn nhầm", () => {
    assert.strictEqual(blockedReason("Giá chuyển nhượng 125000000000 đ theo hợp đồng"), undefined);
  });
});
