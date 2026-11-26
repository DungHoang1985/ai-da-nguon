/** Bước 0 nhóm 2 rà soát kiến trúc (26/11/2026): GOLDEN — đúng thứ add-in gửi model, chụp từ bản 4.13.3.
 *  Mỗi ca chạy một lượt thật qua runTurn với mạng giả, ghi lại yêu cầu ĐẦU TIÊN (địa chỉ, tiêu đề, thân). Thân chứa lời nhắc hệ thống,
 *  danh sách công cụ theo thứ tự (tên, mô tả, tham số), nhóm công cụ (Miễn phí), tham số theo nhà cung cấp — tức mọi thứ quyết định
 *  bộ nhớ đệm và cách model hiểu công cụ. Các bước tách mã sau (B9, B13, B8, B7) phải giữ golden Y NGUYÊN TỪNG BYTE.
 *  Cố ý đổi lời nhắc / công cụ / tham số thì sinh lại: GOLDEN_UPDATE=1 npm test, rồi xem diff của test/fixtures/golden/ trước khi commit. */
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import { runTurn } from "../../src/core/agent/turn";
import { Provider, customProvider, getProvider } from "../../src/core/providers";
import { createWordTools } from "../../src/office/wordTools";
import { excelTools } from "../../src/office/excelTools";
import { pptTools } from "../../src/office/pptTools";
import { createCommonTools } from "../../src/office/common";
import { CORE_GUIDES, BUILTIN_SKILLS } from "../../src/core/skillpack";
import { skillCatalog } from "../../src/core/skills";
import { ToolGroup, toolGroupsFor, visibleSpecs } from "../../src/core/toolgroups";
import { AgentMessage, HostToolSet } from "../../src/core/agent/types";

const DIR = path.join(__dirname, "..", "fixtures", "golden");
const UPDATE = process.env.GOLDEN_UPDATE === "1";
const KEY = "KEY-GIA-golden";

const HOSTS: Record<string, () => HostToolSet> = {
  Word: () => createWordTools(() => ({})),
  Excel: () => excelTools,
  PowerPoint: () => pptTools,
};

/** Chạy một lượt, trả yêu cầu đầu tiên đã chuẩn hóa (không có key thật; thân là JSON đã parse) */
async function capture(o: { provider: string | Provider; model: string; history?: AgentMessage[]; host?: string; mode?: "free" | "paid"; webSearch?: boolean; think?: boolean; pdf?: boolean; noAgent?: boolean; text?: string }) {
  const reqs: { url: string; headers: any; body: any }[] = [];
  const fetchFn = async (url: any, init: any) => {
    const h: Record<string, string> = {};
    const raw = init?.headers || {};
    for (const [k, v] of Object.entries(raw instanceof Headers ? Object.fromEntries((raw as any).entries()) : raw)) h[k.toLowerCase()] = String(v);
    reqs.push({ url: String(url), headers: h, body: init?.body ? JSON.parse(init.body) : null });
    // Dừng lượt ngay sau yêu cầu đầu (401 không thử lại)
    return new Response(JSON.stringify({ error: { message: "golden" } }), { status: 401 });
  };
  const host = o.host;
  const tools = host && !o.noAgent ? HOSTS[host]() : undefined;
  const extra = host && !o.noAgent ? createCommonTools({ host, skills: BUILTIN_SKILLS, shareContext: true, store: new Map() as any, memory: () => [], request: o.text || "Sửa Điều 5." }) : undefined;
  await runTurn({
    provider: typeof o.provider === "string" ? getProvider(o.provider)! : o.provider,
    key: KEY,
    model: o.model,
    host: host || "Word",
    text: o.text || "Sửa Điều 5.",
    history: o.history || [],
    think: !!o.think,
    files: o.pdf ? [{ mime: "application/pdf", data: "JVBERi0xLjQK", name: "thu.pdf" }] : undefined,
    webSearch: !!o.webSearch,
    onEvent: () => {},
    fetchFn,
    retryDelaysMs: [],
    agent: tools
      ? {
          tools,
          extra,
          guide: CORE_GUIDES[host!],
          skillCatalog: skillCatalog(BUILTIN_SKILLS),
          autoSkills: [],
          log: [],
          toolGroups: toolGroupsFor(o.mode || "paid", host!),
          groupHints: [],
          groupSkills: [],
        }
      : undefined,
  } as any);
  assert.ok(reqs.length >= 1, "Không có yêu cầu nào được gửi");
  return reqs[0];
}

function check(name: string, got: any) {
  const file = path.join(DIR, name + ".json");
  const text = JSON.stringify(got, null, 1) + "\n";
  if (UPDATE || !fs.existsSync(file)) {
    if (!UPDATE) assert.fail(`Thiếu golden ${name}.json — chạy GOLDEN_UPDATE=1 npm test để sinh, rồi xem lại trước khi commit`);
    fs.mkdirSync(DIR, { recursive: true });
    fs.writeFileSync(file, text);
    return;
  }
  const want = fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n");
  if (want !== text) {
    // Chỉ ra chỗ lệch đầu tiên cho dễ đọc
    let i = 0;
    while (i < want.length && want[i] === text[i]) i++;
    assert.fail(`Golden ${name}.json lệch ở ký tự ${i}:\n  cũ: …${want.slice(Math.max(0, i - 80), i + 80)}…\n  mới: …${text.slice(Math.max(0, i - 80), i + 80)}…\nCố ý đổi thì GOLDEN_UPDATE=1 npm test và xem diff.`);
  }
}

describe("Golden — thứ add-in gửi model (bước 0 nhóm 2, chụp từ 4.13.3)", function () {
  this.timeout(20000);
  // Ba ứng dụng × Trả phí / Miễn phí: lời nhắc hệ thống + công cụ + nhóm (Miễn phí) — qua nguồn kiểu OpenAI để thân dễ so
  for (const host of Object.keys(HOSTS))
    for (const mode of ["paid", "free"] as const)
      it(`${host} · ${mode === "paid" ? "Trả phí" : "Miễn phí"}`, async () => {
        const r = await capture({ provider: "groq", model: "llama-golden", host, mode });
        check(`app-${host}-${mode}`, r.body);
      });

  // Thân yêu cầu theo nhà cung cấp × tình huống (Word, Trả phí để danh sách công cụ cố định)
  const PROVIDERS: [string, string][] = [
    ["openrouter", "anthropic/claude-sonnet-5.5"],
    ["openai", "gpt-golden"],
    ["groq", "llama-golden"],
    ["mistral", "mistral-golden"],
    ["anthropic", "claude-sonnet-5-5"],
    ["gemini", "gemini-3.8-flash"],
  ];
  const CASES: [string, { webSearch?: boolean; think?: boolean; pdf?: boolean; noAgent?: boolean }][] = [
    ["thuong", {}],
    // Tắt "AI tự thao tác": lời nhắc thường, không công cụ (Gemini chỉ tìm web ở đường này)
    ["khong-tac-tu-web", { noAgent: true, webSearch: true }],
    ["web", { webSearch: true }],
    ["pdf", { pdf: true }],
    ["nghi-ky", { think: true }],
  ];
  // Nguồn kiểu Anthropic có `dialect` (bộ nhớ đệm, cách xác thực, suy nghĩ, PDF, tìm web khác nhau) và nguồn tự khai — đúng phần B9 sẽ đụng
  const CUSTOM_BEARER = customProvider({ id: "tuy-chinh-golden-a", name: "Tự khai A", baseUrl: "https://vi-du.invalid/v1", models: ["m-a"] } as any)!;
  const CUSTOM_XKEY = customProvider({ id: "tuy-chinh-golden-b", name: "Tự khai B", baseUrl: "https://vi-du.invalid/anthropic", anthropicHeaders: true, models: ["m-b"] } as any)!;
  const ALL: [string, string | Provider, string][] = [
    ...PROVIDERS.map(([pid, m]) => [pid, pid, m] as [string, string, string]),
    ["openrouter-claude", "openrouter-claude", "anthropic/claude-sonnet-5.5"],
    ["deepseek-anthropic", "deepseek-anthropic", "deepseek-flash"],
    ["minimax", "minimax", "MiniMax-M3"],
    ["kimi", "kimi", "kimi-k3.1-pro"],
    ["tu-khai-bearer", CUSTOM_BEARER, "m-a"],
    ["tu-khai-xkey", CUSTOM_XKEY, "m-b"],
  ];
  for (const [pid, prov, model] of ALL)
    for (const [cn, c] of CASES)
      it(`${pid} · ${cn}`, async () => {
        const r = await capture({ provider: prov, model, host: "Word", mode: "paid", ...c });
        // Thân đầy đủ đã nằm ở golden app-*: ở đây bỏ lời nhắc và mô tả công cụ dài, giữ cấu trúc + tham số nhà cung cấp + tên công cụ theo thứ tự
        check(`ncc-${pid}-${cn}`, { url: r.url.replace(KEY, "<KEY>"), headers: maskHeaders(r.headers), body: slim(r.body) });
      });

  // Lịch sử có lượt trước (gọi công cụ + kết quả): vị trí đánh dấu bộ nhớ đệm, chuẩn hóa lịch sử theo nguồn đích (có Mistral)
  const HISTORY: AgentMessage[] = [
    { role: "user", content: "Đọc Điều 5." },
    { role: "assistant", content: "", toolCalls: [{ id: "call_golden_1", name: "word_read_clause", args: { clause: "5" } }] } as any,
    { role: "tool", toolCallId: "call_golden_1", name: "word_read_clause", content: '{"ok":true,"text":"Điều 5. Phạt 12%."}' },
    { role: "assistant", content: "Điều 5 quy định phạt 12%." },
  ];
  for (const [pid, prov, model] of ALL.filter(([id]) => ["anthropic", "openai", "mistral", "gemini", "openrouter", "deepseek-anthropic"].includes(id)))
    it(`${pid} · có lịch sử`, async () => {
      const r = await capture({ provider: prov, model, host: "Word", mode: "paid", history: HISTORY, text: "Giảm xuống 8%." });
      check(`lich-su-${pid}`, { url: r.url.replace(KEY, "<KEY>"), body: slim(r.body) });
    });
});

/** Nhóm công cụ (chế độ Miễn phí): công cụ nào hiện ra, theo thứ tự, khi bật từng nhóm / mọi nhóm — B8, B13 không được đổi */
describe("Golden — nhóm công cụ Miễn phí theo từng nhóm", () => {
  for (const host of Object.keys(HOSTS))
    it(host, () => {
      const groups = toolGroupsFor("free", host) as ToolGroup[];
      const specs = HOSTS[host]().specs;
      const one = (active: string[]) => visibleSpecs(specs, groups, new Set(active)).map((t) => t.name);
      check(`nhom-${host}`, {
        nhom: groups.map((g) => ({ id: g.id, label: g.label, tools: g.tools, keywords: g.keywords, skills: g.skills })),
        macDinh: one([]),
        tungNhom: Object.fromEntries(groups.map((g) => [g.id, one([g.id])])),
        tatCa: one(groups.map((g) => g.id)),
      });
    });
});

function maskHeaders(h: Record<string, string>) {
  const out: Record<string, string> = {};
  for (const k of Object.keys(h).sort()) out[k] = h[k].includes(KEY) ? h[k].replace(KEY, "<KEY>") : h[k];
  return out;
}

/** Thu gọn thân: chuỗi dài (lời nhắc, mô tả) thay bằng độ dài + băm ngắn để vẫn bắt được thay đổi mà tệp không phình */
function slim(v: any): any {
  if (typeof v === "string") return v.length > 120 ? `<${v.length} ký tự #${hash(v)}>` : v;
  if (Array.isArray(v)) return v.map(slim);
  if (v && typeof v === "object") return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, slim(x)]));
  return v;
}
function hash(s: string): string {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619);
  return (h >>> 0).toString(16);
}
