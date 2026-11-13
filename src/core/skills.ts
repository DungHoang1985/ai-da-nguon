/**
 * Skills theo chuẩn Agent Skills của Anthropic (SKILL.md: YAML frontmatter name + description, thân là hướng dẫn),
 * nạp dần (progressive disclosure): danh mục tên+mô tả nằm trong lời nhắc hệ thống, thân chỉ đọc khi cần (công cụ use_skill).
 * Lệnh "/" theo chuẩn commands/*.md của plugin Claude ($ARGUMENTS = phần chữ sau tên lệnh).
 */
import { AsyncKV } from "./db";

export interface Skill {
  name: string;
  description: string;
  body: string;
  /** Tệp tham chiếu đi kèm (references/…), chỉ dạng chữ */
  files?: Record<string, string>;
  source: "builtin" | "user";
  /** Nguồn nhập (tên plugin / tệp) */
  origin?: string;
  /** Giới hạn ứng dụng (Excel/Word/PowerPoint); rỗng = mọi ứng dụng */
  hosts?: string[];
  enabled?: boolean;
}

export interface Command {
  name: string;
  description?: string;
  body: string;
  origin?: string;
}

export interface ImportResult {
  skills: Skill[];
  commands: Command[];
  ignored: string[];
}

/** Đọc YAML frontmatter tối giản: key: value, chuỗi trong nháy, khối > hoặc |, một cấp lồng nhau. */
export function parseFrontmatter(text: string): { data: Record<string, any>; body: string } {
  const src = String(text || "").replace(/^﻿/, "");
  const m = src.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/);
  if (!m) return { data: {}, body: src.trim() };
  const lines = m[1].split(/\r?\n/);
  const data: Record<string, any> = {};
  const unquote = (v: string) => {
    const t = v.trim();
    if ((t.startsWith('"') && t.endsWith('"')) || (t.startsWith("'") && t.endsWith("'"))) return t.slice(1, -1);
    return t;
  };
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const kv = line.match(/^([A-Za-z0-9_-]+):\s*(.*)$/);
    if (!kv) continue;
    const key = kv[1];
    const val = kv[2];
    if (val === ">" || val === "|" || val === ">-" || val === "|-") {
      const block: string[] = [];
      while (i + 1 < lines.length && (/^\s+/.test(lines[i + 1]) || lines[i + 1].trim() === "")) block.push(lines[++i].trim());
      data[key] = val.startsWith(">") ? block.join(" ").replace(/\s+/g, " ").trim() : block.join("\n").trim();
    } else if (val === "") {
      const obj: Record<string, any> = {};
      const list: string[] = [];
      while (i + 1 < lines.length && /^\s+/.test(lines[i + 1])) {
        const sub = lines[++i].trim();
        const li = sub.match(/^-\s+(.*)$/);
        const skv = sub.match(/^([A-Za-z0-9_-]+):\s*(.*)$/);
        if (li) list.push(unquote(li[1]));
        else if (skv) obj[skv[1]] = unquote(skv[2]);
      }
      data[key] = list.length ? list : obj;
    } else if (/^\[.*\]$/.test(val.trim())) {
      data[key] = val
        .trim()
        .slice(1, -1)
        .split(",")
        .map((x) => unquote(x))
        .filter(Boolean);
    } else data[key] = unquote(val);
  }
  return { data, body: m[2].trim() };
}

export function toSkillName(s: string): string {
  const n = String(s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 64)
    .replace(/-+$/g, "");
  return n || "ky-nang";
}

export function parseSkillMd(text: string, fallbackName = "", origin?: string): Skill {
  const { data, body } = parseFrontmatter(text);
  const name = toSkillName(data.name || fallbackName);
  const description = String(data.description || "").trim();
  if (!description) throw new Error(`Skill "${name}" thiếu trường description trong frontmatter.`);
  if (!body) throw new Error(`Skill "${name}" không có nội dung hướng dẫn.`);
  const hostsRaw = data.hosts || data.apps || data.metadata?.hosts || data.metadata?.apps;
  const hosts = (Array.isArray(hostsRaw) ? hostsRaw : hostsRaw ? String(hostsRaw).split(",") : [])
    .map((h: string) => normalizeHost(h))
    .filter(Boolean) as string[];
  return { name, description, body, source: "user", origin, hosts: hosts.length ? hosts : undefined, enabled: true };
}

function normalizeHost(h: string): string | null {
  const k = String(h || "").trim().toLowerCase();
  if (k.startsWith("excel")) return "Excel";
  if (k.startsWith("word")) return "Word";
  if (k.startsWith("power") || k === "ppt") return "PowerPoint";
  return null;
}

export function parseCommandMd(text: string, name: string, origin?: string): Command {
  const { data, body } = parseFrontmatter(text);
  if (!body) throw new Error(`Lệnh "${name}" rỗng.`);
  return { name: toSkillName(name), description: data.description ? String(data.description) : undefined, body, origin };
}

const TEXT_EXT = /\.(md|markdown|txt|csv|json|ya?ml|xml|html?)$/i;
const MAX_REF_BYTES = 60000;

/** Nhập gói .zip: một Skill, nhiều Skill, hoặc cả plugin Claude (.claude-plugin/plugin.json + commands/ + skills/). */
export async function importZip(data: ArrayBuffer | Uint8Array, zipName = "goi.zip"): Promise<ImportResult> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(data);
  const paths = Object.keys(zip.files).filter((p) => !zip.files[p].dir && !/(^|\/)(__MACOSX|\.DS_Store)/.test(p));
  const result: ImportResult = { skills: [], commands: [], ignored: [] };
  let origin = zipName.replace(/\.zip$/i, "");
  const manifest = paths.find((p) => /(^|\/)\.claude-plugin\/plugin\.json$/.test(p));
  if (manifest) {
    try {
      const pj = JSON.parse(await zip.files[manifest].async("string"));
      if (pj.name) origin = `plugin ${pj.name}`;
    } catch {
      /* bỏ qua */
    }
  }
  const skillFiles = paths.filter((p) => /(^|\/)SKILL\.md$/i.test(p));
  for (const sf of skillFiles) {
    const dir = sf.slice(0, sf.length - "SKILL.md".length);
    const folderName = dir.replace(/\/$/, "").split("/").pop() || "";
    try {
      const skill = parseSkillMd(await zip.files[sf].async("string"), folderName, origin);
      const files: Record<string, string> = {};
      for (const p of paths) {
        if (p === sf || !p.startsWith(dir)) continue;
        const rel = p.slice(dir.length);
        if (skillFiles.some((o) => o !== sf && o.startsWith(dir) && p.startsWith(o.slice(0, o.length - 8)))) continue;
        if (TEXT_EXT.test(rel)) {
          const content = await zip.files[p].async("string");
          files[rel] = content.length > MAX_REF_BYTES ? content.slice(0, MAX_REF_BYTES) + "\n…[đã cắt]" : content;
        } else result.ignored.push(`${p} (tệp chạy/nhị phân – add-in không thực thi được)`);
      }
      if (Object.keys(files).length) skill.files = files;
      result.skills.push(skill);
    } catch (e: any) {
      result.ignored.push(`${sf}: ${e.message}`);
    }
  }
  for (const p of paths) {
    const cm = p.match(/(^|\/)commands\/(.+)\.md$/i);
    if (cm) {
      try {
        result.commands.push(parseCommandMd(await zip.files[p].async("string"), cm[2].split("/").pop()!, origin));
      } catch (e: any) {
        result.ignored.push(`${p}: ${e.message}`);
      }
    } else if (/(^|\/)(agents\/|hooks\/|\.mcp\.json$)/.test(p)) {
      result.ignored.push(`${p} (agents/hooks/MCP chỉ chạy trong Claude Code)`);
    }
  }
  if (!result.skills.length && !result.commands.length) throw new Error("Gói không có SKILL.md hay commands/*.md nào.");
  return result;
}

export { BUILTIN_SKILLS } from "./skillpack";
import { BUILTIN_SKILLS } from "./skillpack";

const K_SKILLS = "skills-user";
const K_ENABLED = "skills-enabled";
const K_COMMANDS = "commands";

export class SkillStore {
  constructor(private kv: AsyncKV) {}

  async all(): Promise<Skill[]> {
    const user = (await this.kv.get<Skill[]>(K_SKILLS)) || [];
    const enabled = (await this.kv.get<Record<string, boolean>>(K_ENABLED)) || {};
    const byName = new Map<string, Skill>();
    for (const s of BUILTIN_SKILLS) byName.set(s.name, { ...s });
    for (const s of user) byName.set(s.name, { ...s, source: "user" });
    return [...byName.values()].map((s) => ({ ...s, enabled: enabled[s.name] ?? s.enabled ?? true }));
  }

  async forHost(host: string): Promise<Skill[]> {
    return (await this.all()).filter((s) => s.enabled && (!s.hosts?.length || s.hosts.includes(host)));
  }

  async addSkills(skills: Skill[]): Promise<void> {
    const user = (await this.kv.get<Skill[]>(K_SKILLS)) || [];
    const map = new Map(user.map((s) => [s.name, s]));
    for (const s of skills) map.set(s.name, { ...s, source: "user" });
    await this.kv.set(K_SKILLS, [...map.values()]);
  }

  async remove(name: string): Promise<void> {
    const user = ((await this.kv.get<Skill[]>(K_SKILLS)) || []).filter((s) => s.name !== name);
    await this.kv.set(K_SKILLS, user);
  }

  async setEnabled(name: string, on: boolean): Promise<void> {
    const enabled = (await this.kv.get<Record<string, boolean>>(K_ENABLED)) || {};
    enabled[name] = on;
    await this.kv.set(K_ENABLED, enabled);
  }

  async commands(): Promise<Command[]> {
    return (await this.kv.get<Command[]>(K_COMMANDS)) || [];
  }

  async addCommands(cmds: Command[]): Promise<void> {
    const map = new Map((await this.commands()).map((c) => [c.name, c]));
    for (const c of cmds) map.set(c.name, c);
    await this.kv.set(K_COMMANDS, [...map.values()]);
  }

  async removeCommand(name: string): Promise<void> {
    await this.kv.set(
      K_COMMANDS,
      (await this.commands()).filter((c) => c.name !== name)
    );
  }
}

/** Danh mục kỹ năng đưa vào lời nhắc hệ thống (chỉ tên + mô tả). */
export function skillCatalog(skills: Skill[]): string {
  if (!skills.length) return "";
  return [
    "## Kỹ năng (Skills)",
    "Khi yêu cầu khớp mô tả của một kỹ năng, GỌI use_skill với tên kỹ năng để đọc hướng dẫn đầy đủ trước khi làm, rồi làm đúng theo đó. Tệp tham chiếu của kỹ năng đọc bằng use_skill(name, file).",
    ...skills.slice(0, 40).map((s) => `- ${s.name}: ${s.description.slice(0, 300)}`),
  ].join("\n");
}

export interface SlashResult {
  content: string;
  label?: string;
}

/** Xử lý tin nhắn bắt đầu bằng "/tên": lệnh → thay mẫu; skill → nạp sẵn hướng dẫn. */
export function resolveSlash(text: string, skills: Skill[], commands: Command[]): SlashResult {
  const m = String(text || "").match(/^\/([A-Za-z0-9_-]+)\s*([\s\S]*)$/);
  if (!m) return { content: text };
  const name = m[1].toLowerCase();
  const rest = m[2].trim();
  const cmd = commands.find((c) => c.name.toLowerCase() === name);
  if (cmd) {
    const content = /\$ARGUMENTS/.test(cmd.body)
      ? cmd.body.replace(/\$ARGUMENTS/g, rest)
      : rest
        ? `${cmd.body}\n\n${rest}`
        : cmd.body;
    return { content, label: `/${cmd.name}` };
  }
  const sk = skills.find((s) => s.name.toLowerCase() === name);
  if (sk) {
    const files = sk.files ? `\n(Tệp tham chiếu: ${Object.keys(sk.files).join(", ")} — đọc bằng use_skill.)` : "";
    return {
      label: `/${sk.name}`,
      content: `[Người dùng gọi kỹ năng "${sk.name}". Làm đúng theo hướng dẫn của kỹ năng.]\n<skill name="${sk.name}">\n${sk.body}\n</skill>${files}\n\n${rest || "Hãy áp dụng kỹ năng này cho tài liệu đang mở."}`,
    };
  }
  return { content: text };
}

/** Thực thi công cụ use_skill */
export function useSkill(skills: Skill[], name: string, file?: string): any {
  const sk = skills.find((s) => s.name.toLowerCase() === String(name || "").toLowerCase());
  if (!sk) throw new Error(`Không có kỹ năng "${name}". Có: ${skills.map((s) => s.name).join(", ")}`);
  if (file) {
    const content = sk.files?.[file];
    if (content === undefined) throw new Error(`Kỹ năng "${sk.name}" không có tệp "${file}". Có: ${Object.keys(sk.files || {}).join(", ") || "(không)"}`);
    return { ok: true, skill: sk.name, file, content };
  }
  return { ok: true, skill: sk.name, instructions: sk.body, files: Object.keys(sk.files || {}) };
}
