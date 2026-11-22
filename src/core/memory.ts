/** Bộ nhớ dài hạn giữa các hội thoại (Đợt 12, phần E — thiết kế ở tasks/dot-12-bo-nho.md, anh duyệt 23/11/2026).
 *
 *  - Mục ngắn (≤ 300 ký tự) về cách làm việc, thuật ngữ / tên riêng, quy ước văn bản. Chung cho mọi ứng dụng hoặc riêng một ứng dụng.
 *  - Người dùng ghi tay trong Cài đặt; AI chỉ ĐỀ XUẤT (công cụ remember), người dùng bấm Nhớ mới lưu.
 *  - Lưu cục bộ (localStorage của add-in, khóa "aidn.memory"), đi theo bản sao lưu.
 *  - Khi mở hội thoại mới, các mục đang bật được dựng thành một khối chữ, CHỤP LẠI và lưu kèm hội thoại; cả hội thoại dùng đúng
 *    bản chụp đó → lời nhắc hệ thống không đổi giữa các lượt (bộ nhớ đệm lời nhắc giữ nguyên). Mục mới có hiệu lực từ hội thoại sau.
 *  - Chế độ Miễn phí trần thấp hơn: nguồn miễn phí có sức chứa nhỏ và hạn mức token mỗi phút thấp (Groq), lời nhắc cố định đã gần trần. */

export const MEMORY_ITEM_MAX = 300;
/** Trần của khối bộ nhớ trong lời nhắc (ký tự): Trả phí ≈ 550 token, Miễn phí ≈ 220 token */
export const MEMORY_CAP_PAID = 2000;
export const MEMORY_CAP_FREE = 800;
export const MEMORY_MAX_ITEMS = 60;

/** "all" = mọi ứng dụng; còn lại là tên ứng dụng */
export type MemoryScope = "all" | "Word" | "Excel" | "PowerPoint";
export const MEMORY_SCOPES: MemoryScope[] = ["all", "Word", "Excel", "PowerPoint"];

export interface MemoryItem {
  id: string;
  text: string;
  scope: MemoryScope;
  /** Ai ghi: người dùng tự gõ, hoặc AI đề xuất và người dùng bấm Nhớ */
  source: "user" | "ai";
  createdAt: number;
  /** Tắt thì không gửi kèm (vẫn giữ trong danh sách) */
  off?: boolean;
}

export function newMemoryId(now = Date.now()): string {
  return "m" + now.toString(36) + Math.random().toString(36).slice(2, 6);
}

export function normalizeScope(s: unknown, host?: string): MemoryScope {
  const v = String(s ?? "").trim();
  if (v === "this_app" && host && MEMORY_SCOPES.includes(host as MemoryScope)) return host as MemoryScope;
  return MEMORY_SCOPES.includes(v as MemoryScope) ? (v as MemoryScope) : "all";
}

export function cleanMemoryText(t: unknown): string {
  return String(t ?? "")
    .replace(/[\u0000-\u001f\u007f]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** Lý do không được nhớ (chuỗi giống key, số định danh, số điện thoại, số tài khoản); không có thì undefined.
 *  Ưu tiên không chặn nhầm số hiệu văn bản, ngày tháng, số tiền có dấu chấm, số điều khoản. */
export function blockedReason(text: string): string | undefined {
  const t = text.normalize("NFC");
  // Key, token: tiền tố quen của nhà cung cấp, "Bearer …", "api key: …", hoặc chuỗi liền ≥ 32 ký tự lẫn chữ và số
  if (/\b(sk|pk|rk)-[A-Za-z0-9_\-]{12,}/.test(t) || /\bAIza[A-Za-z0-9_\-]{16,}/.test(t) || /\b(gsk|xai|csk|nvapi|pplx|r8|hf)[_-][A-Za-z0-9_\-]{12,}/.test(t))
    return "có chuỗi giống API key";
  if (/\bBearer\s+[A-Za-z0-9._~+/\-]{12,}/i.test(t) || /\b(api[ _-]?key|token|mật khẩu|password|pass)\s*[:=]\s*\S{6,}/i.test(t)) return "có key, token hoặc mật khẩu";
  for (const m of t.match(/[A-Za-z0-9_\-]{32,}/g) || []) if (/[A-Za-z]/.test(m) && /\d/.test(m)) return "có chuỗi giống key hoặc token";
  // Số điện thoại Việt Nam: 0 hoặc +84 rồi 9 chữ số, cho phép cách bằng dấu cách, chấm, gạch
  if (/(?<![\d.,/])(?:\+?84|0)[\s.\-]?[35789](?:[\s.\-]?\d){8}(?![\d,/]|\.\d)/.test(t)) return "có số điện thoại";
  // Dãy ≥ 9 chữ số liền (CMND 9 số, CCCD 12 số, số tài khoản, số thẻ); số tiền viết có dấu chấm ngăn cách không tính
  for (const m of t.match(/(?<![\d.,])\d[\d ]{7,}\d(?![\d,]|\.\d)/g) || []) {
    const digits = m.replace(/ /g, "");
    if (digits.length >= 9 && (!m.includes(" ") || /^\d{4}( \d{4}){3}$/.test(m) || /^\d{3}( \d{3}){3}$/.test(m))) return "có dãy số giống số CCCD, số tài khoản hoặc số thẻ";
  }
  return undefined;
}

/** Kiểm một mục sắp lưu; trả lời lỗi bằng câu tiếng Việt, hợp lệ thì undefined */
export function validateMemoryText(text: string, existing: MemoryItem[] = [], scope?: MemoryScope): string | undefined {
  if (!text) return "Mục nhớ đang trống.";
  if (text.length > MEMORY_ITEM_MAX) return `Mục nhớ dài ${text.length} ký tự, tối đa ${MEMORY_ITEM_MAX}. Hãy viết gọn lại.`;
  const why = blockedReason(text);
  if (why) return `Không nhớ được: ${why}. Bộ nhớ không giữ key, mật khẩu, số định danh, số điện thoại, số tài khoản.`;
  const key = text.toLowerCase();
  if (existing.some((m) => m.text.toLowerCase() === key && (!scope || m.scope === scope))) return "Đã có mục nhớ này.";
  if (existing.length >= MEMORY_MAX_ITEMS) return `Bộ nhớ đã đủ ${MEMORY_MAX_ITEMS} mục. Xóa bớt trong Cài đặt → Bộ nhớ.`;
  return undefined;
}

/** Đọc danh sách đã lưu, bỏ mục hỏng */
export function sanitizeMemory(list: unknown): MemoryItem[] {
  if (!Array.isArray(list)) return [];
  const out: MemoryItem[] = [];
  for (const x of list as any[]) {
    const text = cleanMemoryText(x?.text);
    if (!text || typeof x?.id !== "string") continue;
    out.push({
      id: x.id,
      text: text.slice(0, MEMORY_ITEM_MAX),
      scope: normalizeScope(x.scope),
      source: x.source === "ai" ? "ai" : "user",
      createdAt: Number(x.createdAt) || 0,
      ...(x.off ? { off: true } : {}),
    });
  }
  return out;
}

export interface MemoryBlock {
  /** Khối chữ đưa vào lời nhắc hệ thống ("" khi không có mục nào) */
  text: string;
  /** Số mục đã đưa vào */
  used: number;
  /** Số mục đang bật nhưng không đưa vì vượt trần */
  skipped: number;
}

export const MEMORY_HEADING = "## Bộ nhớ của người dùng (ghi từ các hội thoại trước; áp dụng như chỉ dẫn riêng)";

/** Dựng khối bộ nhớ cho một ứng dụng: mục đang bật của "all" và của ứng dụng đó, theo thứ tự ghi (cũ trước),
 *  dừng khi chạm trần. Cùng đầu vào → cùng chuỗi (không có ngày giờ). */
export function buildMemoryBlock(list: MemoryItem[], host: string, cap: number): MemoryBlock {
  const on = list.filter((m) => !m.off && (m.scope === "all" || m.scope === host)).sort((a, b) => a.createdAt - b.createdAt);
  if (!on.length) return { text: "", used: 0, skipped: 0 };
  const lines: string[] = [];
  let size = MEMORY_HEADING.length;
  let skipped = 0;
  for (const m of on) {
    const line = `- ${m.text}`;
    if (size + 1 + line.length > cap) {
      skipped++;
      continue;
    }
    lines.push(line);
    size += 1 + line.length;
  }
  return { text: lines.length ? [MEMORY_HEADING, ...lines].join("\n") : "", used: lines.length, skipped };
}

export function memoryCapFor(mode?: string): number {
  return mode === "free" ? MEMORY_CAP_FREE : MEMORY_CAP_PAID;
}

/** Đề xuất nhớ do AI đưa ra (chờ người dùng bấm Nhớ / Bỏ) */
export interface MemoryProposal {
  text: string;
  scope: MemoryScope;
  state?: "saved" | "dropped";
  note?: string;
}
