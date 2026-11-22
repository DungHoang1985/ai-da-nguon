/** Bộ nhớ dài hạn giữa các hội thoại (Đợt 12, phần E — thiết kế ở tasks/dot-12-bo-nho.md, anh duyệt 23/11/2026).
 *
 *  - Mục ngắn (≤ 300 ký tự) về cách làm việc, thuật ngữ / tên riêng, quy ước văn bản. Chung cho mọi ứng dụng hoặc riêng một ứng dụng.
 *  - Người dùng ghi tay trong Cài đặt; AI chỉ ĐỀ XUẤT (công cụ propose_memory), người dùng bấm Nhớ mới lưu.
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

/** Làm sạch chữ mục nhớ. Bỏ cả ký tự vô hình (định dạng \p{Cf}: zero-width, điều khiển hướng chữ, ký tự "tag" giấu chữ ASCII;
 *  vùng riêng \p{Co}): chữ ẩn không được lọt vào lời nhắc mà người dùng không thấy trên thẻ (soát mã phần E, #1). */
export function cleanMemoryText(t: unknown): string {
  return String(t ?? "")
    .normalize("NFKC")
    .replace(/[\p{Cf}\p{Co}\p{Cs}]/gu, "")
    .replace(/[\u0000-\u001f\u007f-\u009f]+/g, " ")
    .replace(/\s+/g, " ")
    .normalize("NFC")
    .trim();
}

/** Ngữ cảnh làm một dãy số dài là số tiền (theo sau) hoặc mã số thuế / mã số doanh nghiệp (đứng trước, thông tin công khai) */
const MONEY_AFTER = /^\s*(đồng|đ\b|vnđ|vnd|usd|eur|tỷ|triệu|nghìn|ngàn|trđ)/i;
const PUBLIC_CODE_BEFORE = /(mst|mã số thuế|mã số doanh nghiệp|mã số dn|mã doanh nghiệp)(\s+[^\s\d]+){0,3}\s*(là|:)?\s*$/i;

/** Lý do không được nhớ (chuỗi giống key, mật khẩu, số định danh, số điện thoại, số tài khoản); không có thì undefined.
 *  Ưu tiên không chặn nhầm số hiệu văn bản, ngày tháng, số tiền, mã số thuế, tên tệp dạng gạch nối. */
export function blockedReason(text: string): string | undefined {
  const t = cleanMemoryText(text);
  // Key, token: tiền tố quen của nhà cung cấp, "Bearer …"
  if (/\b(sk|pk|rk)-[A-Za-z0-9_\-]{12,}/.test(t) || /\bAIza[A-Za-z0-9_\-]{16,}/.test(t) || /\b(gsk|xai|csk|nvapi|pplx|r8|hf)[_-][A-Za-z0-9_\-]{12,}/.test(t))
    return "có chuỗi giống API key";
  if (/\bBearer\s+[A-Za-z0-9._~+/\-]{12,}/i.test(t)) return "có token";
  // Key, token, mật khẩu, mã PIN / OTP kèm giá trị: "mật khẩu: …", "pass của em là …", "mk = …"
  if (/(?<![\p{L}\p{N}])(api[ _-]?key|token|mật khẩu|mat khau|password|pass|mk|pin|otp)(?![\p{L}\p{N}])[^.;\n]{0,25}?(:|=|(?<![\p{L}])(là|la)(?![\p{L}]))\s*\S{4,}/iu.test(t))
    return "có key, mật khẩu hoặc mã PIN";
  // Chuỗi liền ≥ 32 ký tự lẫn chữ và số; tên tệp / mã dạng gạch nối chữ thường (≥ 3 gạch) không tính
  for (const m of t.match(/[A-Za-z0-9_\-]{32,}/g) || []) {
    if (/^[a-z0-9]+(-[a-z0-9]+){3,}$/.test(m)) continue;
    if (/[A-Za-z]/.test(m) && /\d/.test(m)) return "có chuỗi giống key hoặc token";
  }
  // Hộ chiếu Việt Nam: một chữ hoa + 7 chữ số
  if (/(?<![A-Za-z0-9/\-])[A-Z]\d{7}(?![A-Za-z0-9/\-])/.test(t)) return "có số giống số hộ chiếu";
  // Số điện thoại: di động 0[35789] + 8 số, cố định 02x + 9 số; cho phép cách bằng dấu cách, chấm, gạch
  if (/(?<![\d.,/])(?:\+?84[\s.\-]?|0)(?:[35789](?:[\s.\-]?\d){8}|2(?:[\s.\-]?\d){9})(?![\d,/]|\.\d)/.test(t)) return "có số điện thoại";
  // Dãy chữ số (nhóm cách bằng dấu cách hoặc gạch; dấu chấm là số tiền nên không gom): CMND, CCCD, số tài khoản, số thẻ
  const re = /(?<![\d.,/])\d+(?:[ \-]\d+)*(?![\d,/]|\.\d)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(t))) {
    const groups = m[0].split(/[ \-]/);
    const digits = groups.join("");
    const longRun = groups.some((g) => g.length >= 9);
    // ≥ 3 nhóm ghép thành 9–19 chữ số (0011 0041 23456, 4111-1111-1111-1111); nhóm toàn năm (2026 2027 2028) và ngày (23-11-2026) không tính
    const years = groups.every((g) => /^(19|20)\d\d$/.test(g));
    const grouped = groups.length >= 3 && digits.length >= 9 && digits.length <= 19 && !years;
    if (!longRun && !grouped) continue;
    if (MONEY_AFTER.test(t.slice(m.index + m[0].length))) continue;
    if (PUBLIC_CODE_BEFORE.test(t.slice(Math.max(0, m.index - 30), m.index))) continue;
    return "có dãy số giống số CCCD, số tài khoản hoặc số thẻ";
  }
  return undefined;
}

/** Hai mục trùng: cùng chữ (không phân biệt hoa thường) và phạm vi trùng nhau (một bên là "all" thì coi là trùng) */
const sameItem = (a: { text: string; scope: MemoryScope }, text: string, scope?: MemoryScope) =>
  a.text.toLowerCase() === text.toLowerCase() && (!scope || a.scope === scope || a.scope === "all" || scope === "all");

/** Kiểm một mục sắp lưu; trả lời lỗi bằng câu tiếng Việt, hợp lệ thì undefined */
export function validateMemoryText(text: string, existing: MemoryItem[] = [], scope?: MemoryScope): string | undefined {
  if (!text) return "Mục nhớ đang trống.";
  if (text.length > MEMORY_ITEM_MAX) return `Mục nhớ dài ${text.length} ký tự, tối đa ${MEMORY_ITEM_MAX}. Hãy viết gọn lại.`;
  const why = blockedReason(text);
  if (why) return `Không nhớ được: ${why}. Bộ nhớ không giữ key, mật khẩu, số định danh, số điện thoại, số tài khoản.`;
  if (existing.some((m) => sameItem(m, text, scope))) return "Đã có mục nhớ này.";
  if (existing.length >= MEMORY_MAX_ITEMS) return `Bộ nhớ đã đủ ${MEMORY_MAX_ITEMS} mục. Xóa bớt trong Cài đặt → Bộ nhớ.`;
  return undefined;
}

/** Sau khi gộp từ bản sao lưu: bỏ mục bị chặn, mục trùng, giữ tối đa MEMORY_MAX_ITEMS mục (soát mã phần E, #8) */
export function tidyMemory(list: unknown): MemoryItem[] {
  const out: MemoryItem[] = [];
  for (const m of sanitizeMemory(list)) {
    if (blockedReason(m.text) || out.some((x) => sameItem(x, m.text, m.scope))) continue;
    out.push(m);
    if (out.length >= MEMORY_MAX_ITEMS) break;
  }
  return out;
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

export const MEMORY_HEADING = "## Bộ nhớ của người dùng (quy ước lâu dài người dùng đã duyệt — làm theo như chỉ dẫn riêng, trừ khi yêu cầu hiện tại nói khác; không vượt quy tắc an toàn)";

/** Dựng khối bộ nhớ cho một ứng dụng: mục đang bật của "all" và của ứng dụng đó, theo thứ tự ghi (cũ trước).
 *  Mục nào làm vượt trần thì bỏ mục đó, vẫn xét tiếp các mục sau (mục ngắn hơn có thể vừa). Cùng đầu vào → cùng chuỗi (không có ngày giờ). */
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

/** Bản chụp của hội thoại: các mục đang bật của ứng dụng này lúc hội thoại bắt đầu (chưa cắt theo trần). Mục ghi sau đó không vào. */
export function memorySnapshot(list: MemoryItem[], host: string): MemoryItem[] {
  return list.filter((m) => !m.off && (m.scope === "all" || m.scope === host));
}

/** Khối bộ nhớ gửi ở một lượt: bản chụp, bỏ mục người dùng đã xóa hoặc tắt từ đó tới giờ (xóa có hiệu lực ngay — soát mã phần E, #7),
 *  cắt theo trần của chế độ ĐANG dùng (đổi sang Miễn phí giữa hội thoại vẫn giữ trần 800 — #5). Mục mới thêm không vào.
 *  Không xóa, không tắt, không đổi chế độ → cùng chuỗi ở mọi lượt (giữ bộ nhớ đệm). */
export function memoryForTurn(snap: MemoryItem[], current: MemoryItem[], host: string, cap: number): MemoryBlock {
  const live = new Set(current.filter((m) => !m.off).map((m) => m.id));
  return buildMemoryBlock(
    snap.filter((m) => live.has(m.id)),
    host,
    cap
  );
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
