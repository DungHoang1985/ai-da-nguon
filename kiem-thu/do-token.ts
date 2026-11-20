/** Đo kích thước phần cố định của một lượt tác tử (lời nhắc hệ thống + bộ công cụ) theo từng ứng dụng.
 *  Chạy: TS_NODE_COMPILER_OPTIONS='{"module":"commonjs","moduleResolution":"node"}' npx ts-node --transpile-only kiem-thu/do-token.ts [Word|Excel|PowerPoint] [--free] [--json]
 *  --free: chế độ Miễn phí (Đợt 10 Word, Đợt 11 Excel, PowerPoint) — chỉ nhóm lõi + load_tools + dòng nhóm công cụ trong lời nhắc.
 *  Số token là ƯỚC LƯỢNG (ký tự / 3,2 cho chữ Việt lẫn Anh); số thật lấy từ usage của nhà cung cấp khi Chạy thật. */
(globalThis as any).Office = (globalThis as any).Office || { context: { requirements: { isSetSupported: () => true } } };
import { getHostTools } from "../src/office/tools";
import { createCommonTools } from "../src/office/common";
import { buildSystemPrompt } from "../src/core/agent/prompt";
import { CORE_GUIDES, BUILTIN_SKILLS } from "../src/core/skillpack";
import { skillCatalog } from "../src/core/skills";
import { freeBasePrompt, groupsNote, hideGroupTools, toolGroupsFor, visibleSpecs, withLoadTools } from "../src/core/toolgroups";

const host = (process.argv[2] || "Word") as any;
const asJson = process.argv.includes("--json");
const groups = process.argv.includes("--free") ? toolGroupsFor("free", host) : undefined;
const free = !!groups;
const tok = (s: string) => Math.round(s.length / 3.2);

const tools = getHostTools(host, () => ({}))!;
const skills = (BUILTIN_SKILLS as any[]).filter((s) => !s.hosts || s.hosts.includes(host));
const common = createCommonTools({ host, skills, shareContext: true });
const extra = groups ? withLoadTools(common, groups) : common;
const parts = {
  base: tools.systemPrompt(),
  guide: CORE_GUIDES[host] || "",
  catalog: skillCatalog(skills),
};
const system = buildSystemPrompt({
  base: groups ? freeBasePrompt(parts.base, groups) : parts.base,
  guide: groups ? hideGroupTools(parts.guide, groups) : parts.guide,
  host,
  skillCatalog: parts.catalog,
  toolGroups: groups ? groupsNote(groups) : undefined,
});
// Dạng gửi Anthropic: { name, description, input_schema }
const specs = groups ? visibleSpecs([...tools.specs, ...extra.specs], groups, new Set()) : [...tools.specs, ...extra.specs];
const toolJson = (t: any) => JSON.stringify({ name: t.name, description: t.description, input_schema: t.parameters || {} });
const perTool = specs
  .map((t) => ({ name: t.name, chars: toolJson(t).length, desc: String(t.description || "").length }))
  .sort((a, b) => b.chars - a.chars);
const toolsChars = perTool.reduce((s, t) => s + t.chars, 0);

const out = {
  host,
  mode: free ? "free (nhóm lõi)" : "full",
  tools: specs.length,
  system_chars: system.length,
  system_parts_chars: {
    base: parts.base.length,
    guide: parts.guide.length,
    catalog: parts.catalog.length,
    rules_and_groups: system.length - parts.base.length - parts.guide.length - parts.catalog.length,
  },
  tools_chars: toolsChars,
  total_chars: system.length + toolsChars,
  est_tokens: { system: tok(system), tools: tok("x".repeat(toolsChars)), total: tok(system) + tok("x".repeat(toolsChars)) },
  per_tool: perTool,
};
if (asJson) console.log(JSON.stringify(out, null, 1));
else {
  console.log(`${host}${free ? " (Miễn phí, nhóm lõi)" : ""}: ${out.tools} công cụ · lời nhắc ${out.system_chars} ký tự · công cụ ${out.tools_chars} ký tự · tổng ≈ ${out.est_tokens.total} token (ước)`);
  console.log("Lời nhắc:", out.system_parts_chars);
  for (const t of perTool) console.log(`  ${t.name.padEnd(30)} ${String(t.chars).padStart(6)}  (mô tả ${t.desc})`);
}
