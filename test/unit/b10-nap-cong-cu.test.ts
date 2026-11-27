/** B10 phần 2 (27/11/2026): bộ công cụ từng ứng dụng nạp riêng bằng import(). Import tĩnh vào đường tải đầu kéo cả bộ công cụ về
 *  gói taskpane mà không báo lỗi gì (phản biện đã bắt một chỗ: Settings.tsx import excelTrail) — chặn bằng máy. */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";

const SRC = path.resolve(__dirname, "../../src");
const HEAVY = /^(excelTools|excelTrail|excelAdvanced|excelModel|wordTools|wordEdit|wordRead|wordComments|wordReview|wordAdmin|wordForm|pptTools|pptAdvanced|pptDesign)$/;

/** Tệp thuộc đường tải đầu của khung: giao diện và tools.ts (selftest* nạp bằng import() nên được phép) */
const entryFiles = () => [
  ...fs.readdirSync(path.join(SRC, "taskpane/components")).map((f) => path.join(SRC, "taskpane/components", f)),
  path.join(SRC, "taskpane/index.tsx"),
  path.join(SRC, "office/tools.ts"),
  path.join(SRC, "office/host.ts"),
  path.join(SRC, "office/navigate.ts"),
];

describe("B10 phần 2 — bộ công cụ không nằm trong gói tải đầu", () => {
  it("không import tĩnh mô-đun công cụ của ứng dụng từ đường tải đầu", () => {
    const bad: string[] = [];
    for (const f of entryFiles()) {
      if (!/\.tsx?$/.test(f) || !fs.existsSync(f)) continue;
      const src = fs.readFileSync(f, "utf8");
      for (const m of src.matchAll(/^\s*import\s+(?!type\b)[^;]*?from\s+"([^"]+)"/gm)) {
        const name = path.basename(m[1]);
        if (HEAVY.test(name)) bad.push(`${path.relative(SRC, f)} → ${m[1]}`);
      }
    }
    assert.deepStrictEqual(bad, [], "Dùng import() (xem loadHostTools trong office/tools.ts): " + bad.join(", "));
  });

  it("loadHostTools nạp cả ba ứng dụng bằng import()", () => {
    const t = fs.readFileSync(path.join(SRC, "office/tools.ts"), "utf8");
    for (const m of ["./excelTools", "./wordTools", "./pptTools"]) assert.ok(t.includes(`"${m}")`), m);
  });
});
