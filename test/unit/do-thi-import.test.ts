/** Nhóm 2 rà soát kiến trúc (26/11/2026): tách tệp (wordCore.ts, các nhóm công cụ Word…) không được sinh import vòng.
 *  Dò mọi import tương đối trong src/ (bỏ `import type` — không chạy lúc chạy thật); có vòng thì nêu đường vòng. */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";

const SRC = path.join(__dirname, "../../src");

function files(dir: string): string[] {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = path.join(dir, e.name);
    return e.isDirectory() ? files(p) : /\.(ts|tsx)$/.test(e.name) && !e.name.endsWith(".d.ts") ? [p] : [];
  });
}

function resolve(from: string, spec: string): string | null {
  const base = path.resolve(path.dirname(from), spec);
  for (const c of [base + ".ts", base + ".tsx", path.join(base, "index.ts"), path.join(base, "index.tsx")]) if (fs.existsSync(c)) return c;
  return null;
}

describe("Đồ thị import của src/ không có vòng", () => {
  it("không có import vòng (bỏ import type, import() động)", () => {
    const graph = new Map<string, string[]>();
    for (const f of files(SRC)) {
      const src = fs.readFileSync(f, "utf8");
      const deps: string[] = [];
      // import … from "./x"  ·  export … from "./x"  (không tính `import type` / `export type`)
      const re = /^\s*(?:import|export)\s+(?!type\b)(?:[^"';]*?\s+from\s+)?["'](\.[^"']+)["']/gm;
      let m: RegExpExecArray | null;
      while ((m = re.exec(src))) {
        const r = resolve(f, m[1]);
        if (r) deps.push(r);
      }
      graph.set(f, deps);
    }
    const state = new Map<string, number>(); // 1 = đang đi, 2 = xong
    const stack: string[] = [];
    const cycles: string[] = [];
    const visit = (n: string) => {
      state.set(n, 1);
      stack.push(n);
      for (const d of graph.get(n) || []) {
        if (state.get(d) === 1) cycles.push([...stack.slice(stack.indexOf(d)), d].map((x) => path.relative(SRC, x)).join(" → "));
        else if (!state.get(d)) visit(d);
      }
      stack.pop();
      state.set(n, 2);
    };
    for (const n of graph.keys()) if (!state.get(n)) visit(n);
    assert.ok(graph.size > 80, `Chỉ thấy ${graph.size} tệp — đường dẫn src/ sai?`);
    assert.deepStrictEqual(cycles, [], "Import vòng:\n" + cycles.join("\n"));
  });
});
