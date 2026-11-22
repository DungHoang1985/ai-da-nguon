/** Sao lưu và khôi phục toàn bộ dữ liệu add-in ra một tệp JSON (Đợt 0A bước 7).
 *  Lý do: Office lưu dữ liệu add-in theo từng tài khoản Office; đổi tài khoản hoặc sau một bản cập nhật Office,
 *  key và lịch sử có thể "biến mất" (thực ra nằm ở kho khác). Tệp sao lưu chứa API key nên là tệp bí mật. */
import type { AsyncKV } from "./db";

export const BACKUP_FORMAT = "ai-da-nguon-sao-luu";
export const BACKUP_FORMAT_VERSION = 1;

/** Phần localStorage cần dùng (giống window.localStorage). */
export interface LocalLike {
  readonly length: number;
  key(i: number): string | null;
  getItem(k: string): string | null;
  setItem(k: string, v: string): void;
  removeItem(k: string): void;
}

export interface BackupFile {
  format: string;
  formatVersion: number;
  appVersion: string;
  createdAt: string;
  host?: string;
  warning: string;
  /** localStorage, tiền tố "aidn." (key, cài đặt, chỉ dẫn riêng, model đã nhớ) — giá trị giữ nguyên chuỗi */
  local: Record<string, string>;
  /** IndexedDB qua db.ts (kỹ năng, lệnh, lịch sử hội thoại) */
  kv: Record<string, unknown>;
}

const PREFIX = "aidn.";
/** Không sao lưu: bộ nhớ đệm danh mục model (tải lại được), dữ liệu đọc chéo (tạm), khóa thử */
const SKIP_LOCAL = (k: string) => k === "aidn.models" || k.startsWith("aidn.share.") || k === "__aidn_test";

const WARNING = "TỆP BÍ MẬT: chứa API key của anh. Không gửi cho người khác, không đưa lên mạng hay vào kho mã.";

export async function createBackup(
  ls: LocalLike,
  kv: AsyncKV,
  meta: { appVersion: string; host?: string; now?: Date }
): Promise<BackupFile> {
  const local: Record<string, string> = {};
  for (let i = 0; i < ls.length; i++) {
    const k = ls.key(i);
    if (!k || !k.startsWith(PREFIX) || SKIP_LOCAL(k)) continue;
    const v = ls.getItem(k);
    if (v !== null) local[k] = v;
  }
  const data: Record<string, unknown> = {};
  for (const k of (await kv.keys("")).sort()) {
    const v = await kv.get(k);
    if (v !== undefined) data[k] = v;
  }
  return {
    format: BACKUP_FORMAT,
    formatVersion: BACKUP_FORMAT_VERSION,
    appVersion: meta.appVersion,
    createdAt: (meta.now || new Date()).toISOString(),
    host: meta.host,
    warning: WARNING,
    local,
    kv: data,
  };
}

export function backupFileName(b: BackupFile): string {
  const d = new Date(b.createdAt);
  const p = (n: number) => String(n).padStart(2, "0");
  return `AI-Da-Nguon-sao-luu-${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}-BI-MAT.json`;
}

export function parseBackup(text: string): BackupFile {
  let j: any;
  try {
    j = JSON.parse(String(text).replace(/^﻿/, ""));
  } catch {
    throw new Error("Không đọc được tệp: nội dung không phải JSON (có thể chép thiếu hoặc tệp hỏng).");
  }
  if (!j || j.format !== BACKUP_FORMAT || typeof j.local !== "object" || typeof j.kv !== "object") {
    throw new Error("Đây không phải tệp sao lưu của AI Đa Nguồn.");
  }
  if (Number(j.formatVersion) > BACKUP_FORMAT_VERSION) {
    throw new Error("Tệp sao lưu do phiên bản mới hơn tạo ra. Hãy cập nhật add-in (Ctrl+F5 trong khung) rồi khôi phục lại.");
  }
  return j as BackupFile;
}

const readJson = (s: string | null): any => {
  if (!s) return undefined;
  try {
    return JSON.parse(s);
  } catch {
    return undefined;
  }
};

export interface BackupSummary {
  keys: number;
  instructions: number;
  conversations: number;
  skills: number;
  commands: number;
  /** Đợt 12 phần E: số mục bộ nhớ */
  memory: number;
}

export function summarizeBackup(b: BackupFile): BackupSummary {
  const keys = readJson(b.local["aidn.keys"]) || {};
  const settings = readJson(b.local["aidn.settings"]) || {};
  const instr = settings.instructions || {};
  return {
    keys: Object.keys(keys).filter((k) => keys[k]).length,
    instructions: Object.keys(instr).filter((k) => String(instr[k] || "").trim()).length,
    conversations: Object.keys(b.kv).filter((k) => k.startsWith("conv:")).length,
    skills: Array.isArray(b.kv["skills-user"]) ? (b.kv["skills-user"] as unknown[]).length : 0,
    commands: Array.isArray(b.kv["commands"]) ? (b.kv["commands"] as unknown[]).length : 0,
    memory: Array.isArray(readJson(b.local["aidn.memory"])) ? readJson(b.local["aidn.memory"]).length : 0,
  };
}

/** Gộp mảng theo khóa: phần tử hiện có được giữ, phần tử chỉ có trong tệp được thêm. */
function unionBy<T>(current: T[], incoming: T[], key: (x: T) => string): T[] {
  const seen = new Set(current.map(key));
  return [...current, ...incoming.filter((x) => !seen.has(key(x)))];
}

/** merge = thêm những gì chưa có, không ghi đè dữ liệu đang có · replace = thay toàn bộ dữ liệu add-in bằng tệp. */
export async function restoreBackup(b: BackupFile, ls: LocalLike, kv: AsyncKV, mode: "merge" | "replace"): Promise<{ local: number; kv: number }> {
  let nLocal = 0;
  let nKv = 0;
  if (mode === "replace") {
    const mine: string[] = [];
    for (let i = 0; i < ls.length; i++) {
      const k = ls.key(i);
      if (k && k.startsWith(PREFIX) && !SKIP_LOCAL(k)) mine.push(k);
    }
    mine.forEach((k) => ls.removeItem(k));
    for (const k of await kv.keys("")) await kv.del(k);
    for (const [k, v] of Object.entries(b.local)) {
      if (!k.startsWith(PREFIX) || SKIP_LOCAL(k)) continue;
      ls.setItem(k, v);
      nLocal++;
    }
    for (const [k, v] of Object.entries(b.kv)) {
      await kv.set(k, v);
      nKv++;
    }
    return { local: nLocal, kv: nKv };
  }

  for (const [k, v] of Object.entries(b.local)) {
    if (!k.startsWith(PREFIX) || SKIP_LOCAL(k)) continue;
    const cur = ls.getItem(k);
    if (cur === null) {
      ls.setItem(k, v);
      nLocal++;
      continue;
    }
    const a = readJson(cur);
    const inc = readJson(v);
    // Bộ nhớ (mảng mục): giữ mục hiện có, thêm mục chỉ có trong tệp (theo mã mục)
    if (k === "aidn.memory" && Array.isArray(a) && Array.isArray(inc)) {
      ls.setItem(k, JSON.stringify(unionBy(a, inc, (x: any) => String(x?.id))));
      nLocal++;
      continue;
    }
    if (a && inc && typeof a === "object" && typeof inc === "object" && !Array.isArray(a)) {
      // Đối tượng (key, cài đặt, lựa chọn): giữ giá trị hiện có, thêm khóa còn thiếu; chỉ dẫn riêng gộp theo từng ứng dụng
      const merged: any = { ...inc, ...a };
      if (inc.instructions || a.instructions) merged.instructions = { ...(inc.instructions || {}), ...(a.instructions || {}) };
      if (inc.models || a.models) merged.models = { ...(inc.models || {}), ...(a.models || {}) };
      for (const [ik, iv] of Object.entries(merged.instructions || {})) if (!String(iv || "").trim() && inc.instructions?.[ik]) merged.instructions[ik] = inc.instructions[ik];
      ls.setItem(k, JSON.stringify(merged));
      nLocal++;
    }
  }

  for (const [k, v] of Object.entries(b.kv)) {
    const cur = await kv.get<any>(k);
    if (cur === undefined) {
      await kv.set(k, v);
      nKv++;
      continue;
    }
    if (k === "conv-index" && Array.isArray(cur) && Array.isArray(v)) {
      const merged = unionBy(cur, v as any[], (x: any) => String(x.id)).sort((x: any, y: any) => (y.updatedAt || 0) - (x.updatedAt || 0));
      await kv.set(k, merged);
      nKv++;
    } else if ((k === "skills-user" || k === "commands") && Array.isArray(cur) && Array.isArray(v)) {
      await kv.set(k, unionBy(cur, v as any[], (x: any) => String(x.name)));
      nKv++;
    } else if (k === "skills-enabled" && cur && typeof cur === "object" && v && typeof v === "object") {
      await kv.set(k, { ...(v as object), ...cur });
      nKv++;
    }
  }
  return { local: nLocal, kv: nKv };
}

/** Kho trống: chưa có key nào và chưa có hội thoại, kỹ năng, lệnh nào. */
export async function looksEmpty(ls: LocalLike, kv: AsyncKV): Promise<boolean> {
  const keys = readJson(ls.getItem("aidn.keys")) || {};
  if (Object.values(keys).some(Boolean)) return false;
  const ks = await kv.keys("");
  return !ks.some((k) => k.startsWith("conv:") || k === "skills-user" || k === "commands");
}
