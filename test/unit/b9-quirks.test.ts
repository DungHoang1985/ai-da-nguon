/** B9 (nhóm 2 rà soát kiến trúc, bước 1, 26/11/2026): đặc điểm riêng của nguồn khai ở providers.ts (`quirks`), lõi không so sánh mã nguồn.
 *  Hành vi y nguyên được giữ bằng golden (golden-gui-model.test.ts); tệp này chặn so sánh mới và giữ các giá trị đã lưu trên máy. */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import { PROVIDERS, customProvider, getProvider, quirksOf } from "../../src/core/providers";
import { cooldownKey, dailyResetAt, nextMidnightIn } from "../../src/core/agent/rotation";
import { callIdOk } from "../../src/core/agent/history";
import { modelInMode, freeKeyWarning } from "../../src/core/modes";
import { pdfCapable, thinkingToggleWorks } from "../../src/core/agent/thinking";

const CORE = path.join(__dirname, "../../src/core");
/** Giao diện cũng quét (phản biện bước 1: 3 chỗ ở Chat, ModelPicker, ProviderCheck đã chuyển sang quirks) */
const UI = path.join(__dirname, "../../src/taskpane");
const NAMES = "openrouter|openai|gemini|groq|mistral|cohere|cerebras|anthropic|deepseek|minimax|kimi|sambanova|nvidia";
/** Tệp được phép so sánh mã nguồn: gắn đúng một nhà cung cấp (API riêng của OpenRouter), hoặc chính nơi khai nguồn */
const ALLOWED: Record<string, string> = {
  "providers.ts": "nơi khai nguồn",
  "credits.ts": "số dư qua API riêng của OpenRouter",
  "oauth.ts": "đăng nhập OAuth của OpenRouter (dạng key sk-or-)",
};

function walk(dir: string): string[] {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walk(path.join(dir, e.name)) : e.name.endsWith(".ts") ? [path.join(dir, e.name)] : []));
}
function walkTsx(dir: string): string[] {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walkTsx(path.join(dir, e.name)) : e.name.endsWith(".tsx") ? [path.join(dir, e.name)] : []));
}

describe("B9 · đặc điểm nguồn (quirks) thay so sánh mã nguồn trong lõi", () => {
  it("lõi không có so sánh `.id === \"…\"` / `providerId === \"…\"` ngoài các tệp được phép", () => {
    const bad: string[] = [];
    for (const f of [...walk(CORE), ...walk(UI).concat(walkTsx(UI))]) {
      const name = path.basename(f);
      if (ALLOWED[name]) continue;
      fs.readFileSync(f, "utf8")
        .split("\n")
        .forEach((line, i) => {
          if (/(\.id|[pP]rovider[iI]d)\s*(===|!==|==|!=)\s*"(?!(string|number|boolean|object|undefined|function)")[a-z][a-z0-9-]*"/.test(line)) bad.push(`${path.relative(CORE, f)}:${i + 1}: ${line.trim()}`);
          if (/new Set\(\s*\[\s*"(openrouter|openai|gemini|groq|mistral|cohere|cerebras|anthropic)"/.test(line)) bad.push(`${path.relative(CORE, f)}:${i + 1}: tập mã nguồn — khai thành quirks`);
          // Các dạng khác (phản biện bước 1): "x" === p.id, case "x", biểu thức chính quy /^openrouter…/ trên mã nguồn, [..."x"...].includes(…id)
          if (new RegExp(String.raw`"(${NAMES})(-[a-z]+)?"\s*(===|!==)\s*[\w.?]*([iI]d)\b`).test(line)) bad.push(`${path.relative(CORE, f)}:${i + 1}: ${line.trim()}`);
          if (new RegExp(String.raw`case\s+"(${NAMES})(-[a-z]+)?"\s*:`).test(line)) bad.push(`${path.relative(CORE, f)}:${i + 1}: ${line.trim()}`);
          if (new RegExp(String.raw`/\^(${NAMES})`).test(line) && /[iI]d\)/.test(line)) bad.push(`${path.relative(CORE, f)}:${i + 1}: ${line.trim()}`);
          if (new RegExp(String.raw`\[\s*"(${NAMES})"[^\]]*\]\.includes\([\w.?]*[iI]d\)`).test(line)) bad.push(`${path.relative(CORE, f)}:${i + 1}: ${line.trim()}`);
        });
    }
    assert.deepStrictEqual(bad, [], "Khai đặc điểm vào `quirks` ở providers.ts thay vì so sánh mã nguồn:\n" + bad.join("\n"));
  });

  it("khóa thời gian nghỉ đã lưu trên máy không đổi giá trị", () => {
    assert.strictEqual(cooldownKey({ provider: { id: "openrouter" }, model: "x:free" }, "daily"), "openrouter");
    assert.strictEqual(cooldownKey({ provider: getProvider("openrouter")!, model: "x:free" }, "daily"), "openrouter");
    assert.strictEqual(cooldownKey({ provider: { id: "openrouter" }, model: "x:free" }, "rate" as any), "openrouter|x:free");
    assert.strictEqual(cooldownKey({ provider: { id: "gemini" }, model: "g" }, "daily"), "gemini|g");
    assert.strictEqual(cooldownKey({ provider: { id: "groq" }, model: "m" }), "groq|m");
  });

  it("mốc đặt lại hạn mức ngày giữ như cũ: Gemini nửa đêm giờ Thái Bình Dương, OpenRouter nửa đêm UTC, nguồn khác 24 giờ", () => {
    const now = Date.UTC(2026, 10, 26, 3, 0, 0);
    assert.strictEqual(dailyResetAt("openrouter", now), Date.UTC(2026, 10, 27, 0, 0, 0));
    assert.strictEqual(dailyResetAt("groq", now), now + 86400000);
    assert.strictEqual(dailyResetAt("gemini", now), nextMidnightIn("America/Los_Angeles", now));
  });

  it("nguồn tự khai không có đặc điểm nào; nguồn lạ cũng vậy", () => {
    const c = customProvider({ id: "tuy-chinh-q", name: "Q", baseUrl: "https://vi-du.invalid/v1", models: ["m"] } as any)!;
    assert.deepStrictEqual(quirksOf(c), {});
    assert.deepStrictEqual(quirksOf({ id: "khong-co" }), {});
    assert.deepStrictEqual(quirksOf(undefined), {});
  });

  it("đặc điểm của từng nguồn đúng bảng kỳ vọng (mọi trường, mọi nguồn)", () => {
    const want: Record<string, object> = {
      openrouter: { reasoning: "openrouter", reasoningDetails: true, appTitleHeader: true, webPlugin: true, pdf: true, pdfPlugin: true, callId: "dotted", catalogPricing: true, mixedFreeCatalog: true, accountDailyQuota: true, dailyResetTz: "UTC", freeKeyNote: "none", relay: true },
      "openrouter-claude": { catalogPricing: true, relay: true },
      openai: { reasoning: "openai", pdf: true },
      mistral: { callId: "nine", assistantAfterTool: true, ambiguousFreeTier: true },
      gemini: { freeByPattern: true, dailyResetTz: "America/Los_Angeles", ambiguousFreeTier: true },
      cohere: { freeKeyNote: "trial", ambiguousFreeTier: true },
      groq: { ambiguousFreeTier: true },
      cerebras: { ambiguousFreeTier: true },
    };
    const got = Object.fromEntries(PROVIDERS.filter((p) => p.quirks).map((p) => [p.id, p.quirks]));
    assert.deepStrictEqual(got, want);
  });

  it("hành vi theo nguồn giữ như cũ (mẫu)", () => {
    const or = getProvider("openrouter")!, oai = getProvider("openai")!, mis = getProvider("mistral")!, gq = getProvider("groq")!, orc = getProvider("openrouter-claude")!;
    assert.ok(callIdOk(mis, "abcDEF123") && !callIdOk(mis, "call_1"));
    assert.ok(callIdOk(or, "functions.ten:0") && !callIdOk(gq, "functions.ten:0"));
    assert.ok(thinkingToggleWorks(or, "x") && thinkingToggleWorks(oai, "x") && !thinkingToggleWorks(gq, "x"));
    assert.ok(pdfCapable(or) && pdfCapable(oai) && !pdfCapable(gq));
    assert.strictEqual(modelInMode(or, { id: "a:free", free: true }, "paid"), false);
    assert.strictEqual(modelInMode(orc, { id: "a", free: true }, "paid"), modelInMode(orc, { id: "a" }, "paid"));
    assert.strictEqual(freeKeyWarning(or), "");
    assert.match(freeKeyWarning(getProvider("cohere")!), /^Key dùng thử/);
  });
});
