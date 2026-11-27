/** Nhóm 4 rà soát kiến trúc (bảo mật nhỏ, 27/11/2026): B11, B17, B19, B20, B21, B22 */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import { execSync } from "child_process";
import { FallbackKV, IdbKV, MemoryKV } from "../../src/core/db";
import { ConversationStore } from "../../src/core/conversations";
import { Storage } from "../../src/core/storage";
import { allProviders, CUSTOM_PREFIX, setCustomSources } from "../../src/core/providers";
// eslint-disable-next-line @typescript-eslint/no-require-imports
const { buildCsp, injectCsp } = require("../../csp-plugin");

const ROOT = path.resolve(__dirname, "../..");
const conv = (id: string, host = "Word") => ({ id, host, title: id, updatedAt: 0, items: [], history: [] });

describe("Nhóm 4 — B11 kho lưu trên máy", () => {
  it("update đọc – sửa – ghi một lần, trả giá trị đã ghi", async () => {
    const kv = new MemoryKV();
    await kv.set("n", 1);
    assert.strictEqual(await kv.update<number>("n", (v) => (v || 0) + 1), 2);
    assert.strictEqual(await kv.get("n"), 2);
  });

  it("hai khung cùng lưu: mục lục giữ đủ cả hai hội thoại", async () => {
    const kv = new MemoryKV();
    const a = new ConversationStore(kv);
    const b = new ConversationStore(kv);
    await Promise.all([a.save(conv("x")), b.save(conv("y", "Excel")), a.save(conv("z"))]);
    const ids = (await a.list()).map((c) => c.id).sort();
    assert.deepStrictEqual(ids, ["x", "y", "z"]);
  });

  it("xóa theo ứng dụng chỉ xóa đúng ứng dụng đó", async () => {
    const kv = new MemoryKV();
    const st = new ConversationStore(kv);
    await st.save(conv("w1"));
    await st.save(conv("e1", "Excel"));
    await st.clear("Word");
    assert.deepStrictEqual((await st.list()).map((c) => c.id), ["e1"]);
    assert.strictEqual(await kv.get("conv:w1"), undefined);
  });

  it("mở IndexedDB lỗi: chuyển sang bộ nhớ tạm, không hỏng mãi, và báo", async () => {
    const fake = Object.create(IdbKV.prototype) as IdbKV;
    (fake as any).ready = Promise.reject(new Error("hỏng"));
    (fake as any).ready.catch(() => undefined);
    for (const m of ["get", "set", "del", "keys", "update"]) (fake as any)[m] = () => fake.ready;
    const kv = new FallbackKV(() => fake);
    const seen: string[] = [];
    kv.onFallback((w) => seen.push(w));
    await kv.set("a", 1);
    assert.strictEqual(await kv.get("a"), 1);
    assert.strictEqual(kv.degraded, true);
    assert.deepStrictEqual(seen, ["hỏng"]);
    // Đăng ký sau khi đã chuyển vẫn được báo
    const late: string[] = [];
    kv.onFallback((w) => late.push(w));
    assert.deepStrictEqual(late, ["hỏng"]);
  });

  it("không có IndexedDB: dùng bộ nhớ tạm và báo", async () => {
    const kv = new FallbackKV(() => {
      throw new Error("Không có IndexedDB");
    });
    assert.strictEqual(kv.degraded, true);
    await kv.set("k", "v");
    assert.strictEqual(await kv.get("k"), "v");
  });

  it("IdbKV báo xong ở oncomplete của giao dịch, không ở onsuccess của lệnh", () => {
    const src = fs.readFileSync(path.join(ROOT, "src/core/db.ts"), "utf8");
    const body = src.slice(src.indexOf("class IdbKV"), src.indexOf("class FallbackKV"));
    assert.match(body, /t\.oncomplete = \(\) => resolve/);
    assert.doesNotMatch(body, /r\.onsuccess = \(\) => resolve/);
  });
});

describe("Nhóm 4 — B22 nguồn tự khai nạp tường minh", () => {
  it("tạo Storage là có nguồn tự khai trong allProviders(), không cần gọi getKeys()", () => {
    setCustomSources([]);
    const mem = new Map<string, string>();
    const store = { getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v), removeItem: (k: string) => void mem.delete(k) };
    new Storage(store as any).setCustomSources([{ id: CUSTOM_PREFIX + "b22", name: "Nguồn thử B22", baseUrl: "https://vd.example/anthropic", models: ["m1"] }]);
    setCustomSources([]);
    assert.ok(!allProviders().some((p) => p.name === "Nguồn thử B22"));
    new Storage(store as any);
    assert.ok(allProviders().some((p) => p.name === "Nguồn thử B22"));
    setCustomSources([]);
  });

  it("getKeys() không còn tác dụng phụ đăng ký nguồn", () => {
    const src = fs.readFileSync(path.join(ROOT, "src/core/storage.ts"), "utf8");
    const body = src.slice(src.indexOf("getKeys(): Record"), src.indexOf("getCustomSources(): CustomSource"));
    assert.doesNotMatch(body, /registerCustomSources/);
  });
});

describe("Nhóm 4 — B19 Word dưới 1.6 không ghi", () => {
  it("tracked() đòi WordApi 1.6 (Hoàn tác cần đọc danh sách sửa đổi)", () => {
    const src = fs.readFileSync(path.join(ROOT, "src/office/wordCore.ts"), "utf8");
    const body = src.slice(src.indexOf("export async function tracked"), src.indexOf("export async function targetRanges"));
    assert.match(body, /if \(!supports\("1\.6"\)\)/);
  });
});

describe("Nhóm 4 — B17 CSP", () => {
  it("mã băm của script nội tuyến tính trên HTML đã rút gọn, script có src thì bỏ qua", () => {
    const html = `<head><meta charset="UTF-8"/><script>try{a()}catch(t){}</script><script src="https://officeapis.public.onecdn.static.microsoft/1/office.js"></script></head>`;
    const csp = buildCsp(html);
    const hash = require("crypto").createHash("sha256").update("try{a()}catch(t){}", "utf8").digest("base64");
    assert.ok(csp.includes(`'sha256-${hash}'`));
    assert.strictEqual((csp.match(/sha256-/g) || []).length, 1);
    assert.match(csp, /object-src 'none'/);
    assert.match(csp, /base-uri 'self'/);
    assert.doesNotMatch(csp, /script-src[^;]*'unsafe-(inline|eval)'/);
    const out = injectCsp(html);
    assert.ok(out.indexOf("Content-Security-Policy") < out.indexOf("<script"));
    assert.strictEqual(injectCsp(out), out);
  });

  it("trang thiếu <meta charset> thì dừng build, không lặng lẽ bỏ CSP", () => {
    assert.throws(() => injectCsp("<head><script>x()</script></head>"), /meta charset/);
  });

  it("webpack dùng CspPlugin", () => {
    assert.match(fs.readFileSync(path.join(ROOT, "webpack.config.js"), "utf8"), /new CspPlugin\(/);
  });
});

describe("Nhóm 4 — B21 cổng gỡ lỗi khi mở Office thử", () => {
  it("không cho mọi nguồn (remote-allow-origins=*)", () => {
    const src = fs.readFileSync(path.join(ROOT, "kiem-thu/mo-office.ps1"), "utf8");
    assert.doesNotMatch(src, /remote-allow-origins=\*/);
    assert.match(src, /remote-allow-origins=http:\/\/127\.0\.0\.1:/);
  });
});

describe("Nhóm 4 — B20 kho công khai không chứa đường dẫn máy", () => {
  it("không tệp nào trong Git có đường dẫn tuyệt đối kiểu ổ đĩa Windows tới thư mục người dùng hoặc dự án", function () {
    let files: string[];
    try {
      files = execSync("git ls-files -z", { cwd: ROOT, encoding: "utf8", maxBuffer: 64 << 20 }).split("\0").filter(Boolean);
    } catch {
      this.skip();
      return;
    }
    // Ghép mẫu từng mảnh để chính tệp này không khớp
    const re = new RegExp("[A-Za-z]:" + "\\\\{1,2}" + "(Users|ADD IN|CLAUDE)" + "\\\\{1,2}", "i");
    const bad: string[] = [];
    for (const f of files) {
      if (/\.(png|jpe?g|gif|ico|pdf|docx|xlsx|pptx|zip|woff2?|ttf)$/i.test(f)) continue;
      let text: string;
      try {
        text = fs.readFileSync(path.join(ROOT, f), "utf8");
      } catch {
        continue;
      }
      if (re.test(text)) bad.push(f);
    }
    assert.deepStrictEqual(bad, [], "Đổi sang đường dẫn tương đối hoặc <worktree>: " + bad.join(", "));
  });
});
