/** Đợt 0A – bước 7: sao lưu và khôi phục toàn bộ dữ liệu add-in ra một tệp. */
import * as assert from "assert";
import { MemoryKV } from "../../src/core/db";
import { BACKUP_FORMAT, LocalLike, createBackup, looksEmpty, parseBackup, restoreBackup, summarizeBackup } from "../../src/core/backup";

function fakeLocal(init: Record<string, string> = {}): LocalLike & { d: Record<string, string> } {
  const d: Record<string, string> = { ...init };
  return {
    d,
    get length() {
      return Object.keys(d).length;
    },
    key: (i: number) => Object.keys(d)[i] ?? null,
    getItem: (k: string) => (k in d ? d[k] : null),
    setItem: (k: string, v: string) => {
      d[k] = v;
    },
    removeItem: (k: string) => {
      delete d[k];
    },
  };
}

async function sample() {
  const ls = fakeLocal({
    "aidn.keys": JSON.stringify({ openrouter: "sk-or-1", anthropic: "sk-ant-1" }),
    "aidn.settings": JSON.stringify({ askBeforeWrite: true, instructions: { Word: "Văn phong hành chính" } }),
    "aidn.selection": JSON.stringify({ providerId: "openrouter", models: { openrouter: "a/b:free" } }),
    "aidn.models": JSON.stringify({ openrouter: { at: 1, models: [] } }),
    "aidn.share.Excel": "{}",
    "ung-dung-khac": "x",
  });
  const kv = new MemoryKV();
  await kv.set("skills-user", [{ name: "ky-nang-thu", description: "d", body: "b", source: "user" }]);
  await kv.set("skills-enabled", { "ky-nang-thu": true });
  await kv.set("commands", [{ name: "tom-tat", body: "Tóm tắt $ARGUMENTS" }]);
  await kv.set("conv-index", [{ id: "c1", host: "Word", title: "Hội thoại 1", updatedAt: 10 }]);
  await kv.set("conv:c1", { id: "c1", host: "Word", title: "Hội thoại 1", updatedAt: 10, items: [{ role: "user", text: "xin chào" }], history: [] });
  return { ls, kv };
}

describe("Đợt 0A – bước 7: sao lưu và khôi phục", () => {
  it("Sao lưu gom key, cài đặt, chỉ dẫn riêng, model đã nhớ, kỹ năng, lệnh, lịch sử; bỏ bộ nhớ đệm danh mục và dữ liệu đọc chéo", async () => {
    const { ls, kv } = await sample();
    const b = await createBackup(ls, kv, { appVersion: "3.6.0", host: "Word", now: new Date("2026-11-15T03:00:00Z") });
    assert.strictEqual(b.format, BACKUP_FORMAT);
    assert.strictEqual(b.formatVersion, 1);
    assert.match(b.warning, /BÍ MẬT/);
    assert.deepStrictEqual(Object.keys(b.local).sort(), ["aidn.keys", "aidn.selection", "aidn.settings"]);
    assert.deepStrictEqual(Object.keys(b.kv).sort(), ["commands", "conv-index", "conv:c1", "skills-enabled", "skills-user"]);
    const sum = summarizeBackup(b);
    assert.strictEqual(sum.keys, 2);
    assert.strictEqual(sum.conversations, 1);
    assert.strictEqual(sum.skills, 1);
    assert.strictEqual(sum.commands, 1);
    assert.strictEqual(sum.instructions, 1);
  });

  it("Đọc tệp: sai định dạng hoặc hỏng → báo lỗi tiếng Việt", () => {
    assert.throws(() => parseBackup("không phải json"), /không đọc được/i);
    assert.throws(() => parseBackup(JSON.stringify({ a: 1 })), /không phải tệp sao lưu/i);
    assert.throws(() => parseBackup(JSON.stringify({ format: BACKUP_FORMAT, formatVersion: 99, local: {}, kv: {} })), /phiên bản mới hơn/i);
  });

  it("Sao lưu → xóa dữ liệu thử → khôi phục (gộp): dữ liệu trở lại đủ, dữ liệu hiện có được giữ", async () => {
    const { ls, kv } = await sample();
    const text = JSON.stringify(await createBackup(ls, kv, { appVersion: "3.6.0", host: "Word" }));
    // xóa dữ liệu thử
    ls.setItem("aidn.keys", JSON.stringify({ groq: "gsk-moi" }));
    ls.setItem("aidn.settings", JSON.stringify({ askBeforeWrite: false, instructions: {} }));
    await kv.del("conv:c1");
    await kv.set("conv-index", [{ id: "c2", host: "Excel", title: "Mới", updatedAt: 20 }]);
    await kv.set("conv:c2", { id: "c2" });
    await kv.set("skills-user", []);
    const r = await restoreBackup(parseBackup(text), ls, kv, "merge");
    assert.ok(r.local > 0 && r.kv > 0);
    assert.deepStrictEqual(JSON.parse(ls.getItem("aidn.keys")!), { openrouter: "sk-or-1", anthropic: "sk-ant-1", groq: "gsk-moi" });
    const st = JSON.parse(ls.getItem("aidn.settings")!);
    assert.strictEqual(st.askBeforeWrite, false, "giữ cài đặt hiện có");
    assert.strictEqual(st.instructions.Word, "Văn phong hành chính", "chỉ dẫn riêng trở lại");
    const idx = (await kv.get<any[]>("conv-index"))!;
    assert.deepStrictEqual(idx.map((x) => x.id), ["c2", "c1"]);
    assert.ok(await kv.get("conv:c1"));
    assert.ok(await kv.get("conv:c2"));
    assert.strictEqual(((await kv.get<any[]>("skills-user")) || []).length, 1);
  });

  it("Khôi phục (thay toàn bộ): dữ liệu add-in hiện có bị thay đúng bằng tệp; không đụng dữ liệu ngoài add-in", async () => {
    const { ls, kv } = await sample();
    const b = await createBackup(ls, kv, { appVersion: "3.6.0", host: "Word" });
    ls.setItem("aidn.keys", JSON.stringify({ groq: "gsk-moi" }));
    await kv.set("conv-index", [{ id: "c2", host: "Excel", title: "Mới", updatedAt: 20 }]);
    await kv.set("conv:c2", { id: "c2" });
    await restoreBackup(b, ls, kv, "replace");
    assert.deepStrictEqual(JSON.parse(ls.getItem("aidn.keys")!), { openrouter: "sk-or-1", anthropic: "sk-ant-1" });
    assert.strictEqual(await kv.get("conv:c2"), undefined);
    assert.strictEqual(ls.getItem("ung-dung-khac"), "x");
  });

  it("Kho trống (dấu hiệu đổi tài khoản Office hoặc Office đổi thư mục lưu) → nhận ra để nhắc khôi phục", async () => {
    assert.strictEqual(await looksEmpty(fakeLocal({ "aidn.settings": "{}" }), new MemoryKV()), true);
    const { ls, kv } = await sample();
    assert.strictEqual(await looksEmpty(ls, kv), false);
  });
});
