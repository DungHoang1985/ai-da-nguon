/** Ghi chéo ứng dụng (Đợt 3, việc B1–B2): khung ở ứng dụng này nhờ khung AI Đa Nguồn ở ứng dụng kia ghi hộ vào tài liệu của nó.
 *  Không có máy chủ: "hộp thư" là các khóa trong localStorage mà ba ứng dụng Office dùng chung (đã đo 16/11/2026: sự kiện storage
 *  tới khung kia trong 0–2 ms, khung ở ứng dụng nền vẫn ghi được, hẹn giờ không bị bóp — bảng ở tasks/dot/dot-3/dot-3.md).
 *
 *  An toàn:
 *  - Người dùng LUÔN được hỏi ở khung gửi trước khi lệnh rời đi (công cụ write_other_app có cảnh báo rủi ro ở mọi lần gọi;
 *    "Cho phép hết lần này" không áp dụng cho thao tác rủi ro).
 *  - Khung nhận chỉ chạy thao tác trong một danh sách đóng (thêm nội dung mới, không sửa hay xóa cái có sẵn), kiểm lại hình dạng và
 *    kích thước của lệnh, bỏ lệnh quá hạn hoặc đã chạy, và chỉ nhận khi công tắc "nhận ghi từ ứng dụng khác" đang bật.
 *  - Nội dung của lệnh là dữ liệu để chèn, không phải chỉ dẫn: khung nhận không đưa nó cho model.
 *  Mọi hàm ở đây thuần (nhận kho và đồng hồ từ ngoài) để kiểm thử được ngoài Office. */
import { KVStore } from "./storage";

export type AppName = "Excel" | "Word" | "PowerPoint";
export const APPS: AppName[] = ["Excel", "Word", "PowerPoint"];

export const LIVE_PREFIX = "aidn.live.";
export const CMD_PREFIX = "aidn.cmd.";
export const RES_PREFIX = "aidn.res.";
/** Nhịp sống: mỗi khung ghi lại mỗi 3 giây; quá 10 giây không thấy thì coi là khung đã đóng */
export const HEARTBEAT_MS = 3000;
export const LIVE_MAX_AGE_MS = 10000;
/** Lệnh chưa được nhận sau bấy lâu thì bỏ (khung nhận không chạy lệnh cũ). Dài hơn hạn chờ của khung gửi: lệnh còn hạn mà khung gửi
 *  đã thôi chờ thì chính khung gửi rút lệnh, nên "đã nhận mà chưa báo kết quả" luôn nghĩa là khung nhận đã lấy thật. */
export const CMD_MAX_AGE_MS = 30000;
/** Khung gửi chờ kết quả tối đa bấy lâu */
export const SEND_TIMEOUT_MS = 25000;
export const MAX_TEXT_CHARS = 20000;
export const MAX_CELLS = 5000;

export interface Live {
  host: AppName;
  at: number;
  /** Tên tệp đang mở (để người dùng biết sẽ ghi vào đâu) */
  title?: string;
  /** Khung này có nhận lệnh ghi từ ứng dụng khác không */
  accept: boolean;
}

/** Thao tác ghi hộ — danh sách đóng, mỗi ứng dụng đúng một thao tác, chỉ THÊM nội dung mới */
export type CrossOp =
  | { op: "word_append"; text: string }
  | { op: "excel_new_sheet"; name: string; values: string[][] }
  | { op: "ppt_add_slide"; title: string; body: string };

export const OP_OF: Record<AppName, CrossOp["op"]> = { Word: "word_append", Excel: "excel_new_sheet", PowerPoint: "ppt_add_slide" };

export interface Command {
  id: string;
  from: AppName;
  to: AppName;
  at: number;
  /** Tên tệp đích mà người dùng đã thấy trong hộp hỏi. Mở hai tài liệu cùng ứng dụng thì chỉ khung của đúng tệp này nhận lệnh. */
  toTitle?: string;
  payload: CrossOp;
}

export interface CommandResult {
  id: string;
  ok: boolean;
  at: number;
  message?: string;
  error?: string;
}

const isApp = (v: any): v is AppName => APPS.includes(v);
const read = <T>(s: KVStore, key: string): T | null => {
  try {
    const raw = s.getItem(key);
    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
};
const write = (s: KVStore, key: string, v: unknown): boolean => {
  try {
    s.setItem(key, JSON.stringify(v));
    return true;
  } catch {
    return false;
  }
};

export function normalizeApp(app: any): AppName | null {
  const k = String(app || "")
    .trim()
    .toLowerCase();
  if (k.startsWith("excel")) return "Excel";
  if (k.startsWith("word")) return "Word";
  if (k.startsWith("power") || k === "ppt") return "PowerPoint";
  return null;
}

// ---------- Nhịp sống ----------

export function beat(s: KVStore, live: Omit<Live, "at">, now = Date.now()): void {
  write(s, LIVE_PREFIX + live.host, { ...live, title: live.title ? String(live.title).slice(0, 120) : undefined, at: now });
}

export function clearBeat(s: KVStore, host: AppName): void {
  try {
    s.removeItem(LIVE_PREFIX + host);
  } catch {
    /* bỏ qua */
  }
}

/** Khung ở ứng dụng `host` có đang mở không (theo nhịp sống) */
export function liveOf(s: KVStore, host: AppName, now = Date.now()): Live | null {
  const v = read<Live>(s, LIVE_PREFIX + host);
  if (!v || v.host !== host || typeof v.at !== "number" || now - v.at > LIVE_MAX_AGE_MS || v.at - now > LIVE_MAX_AGE_MS) return null;
  return { host, at: v.at, title: typeof v.title === "string" ? v.title : undefined, accept: v.accept === true };
}

export interface Connection {
  app: AppName;
  /** Khung ở ứng dụng đó đang mở */
  open: boolean;
  /** Khung đó nhận lệnh ghi */
  accept: boolean;
  title?: string;
}

/** Trạng thái kết nối tới các ứng dụng khác, cho chỉ báo trên khung */
export function connections(s: KVStore, self: AppName, now = Date.now()): Connection[] {
  return APPS.filter((a) => a !== self).map((app) => {
    const l = liveOf(s, app, now);
    return { app, open: !!l, accept: !!l?.accept, title: l?.title };
  });
}

/** Dòng chỉ báo ngắn, ví dụ "Excel ✓ · PowerPoint —" */
export function connectionLine(cs: Connection[]): string {
  return cs.map((c) => `${c.app} ${c.open ? "✓" : "—"}`).join(" · ");
}

// ---------- Kiểm lệnh ----------

/** Kiểm và chuẩn hóa nội dung ghi hộ cho ứng dụng đích. Ném lỗi tiếng Việt khi không hợp lệ. */
export function buildPayload(to: AppName, a: any): CrossOp {
  const text = (v: any) => (typeof v === "string" ? v : Array.isArray(v) ? v.map((x) => String(x ?? "")).join("\n") : v === undefined || v === null ? "" : String(v));
  if (to === "Word") {
    const t = text(a?.text).trim();
    if (!t) throw new Error("Thiếu text: nội dung cần thêm vào cuối văn bản Word (mỗi dòng một đoạn).");
    if (t.length > MAX_TEXT_CHARS) throw new Error(`Nội dung quá dài (${t.length} ký tự, tối đa ${MAX_TEXT_CHARS}): chia thành nhiều lần ghi.`);
    return { op: "word_append", text: t };
  }
  if (to === "Excel") {
    const rows = Array.isArray(a?.values) ? a.values : null;
    if (!rows || !rows.length || !rows.every((r: any) => Array.isArray(r))) throw new Error("Thiếu values: bảng dữ liệu dạng mảng các hàng, ví dụ [[\"Hạng mục\",\"Giá trị\"],[\"A\",\"1\"]].");
    const width = Math.max(...rows.map((r: any[]) => r.length));
    if (!width) throw new Error("values rỗng.");
    if (rows.length * width > MAX_CELLS) throw new Error(`Bảng quá lớn (${rows.length}×${width} ô, tối đa ${MAX_CELLS} ô).`);
    // Không nhận công thức từ ứng dụng khác: ô bắt đầu bằng =, +, -, @ được ghi thành chữ (chống chèn công thức)
    const values = rows.map((r: any[]) => Array.from({ length: width }, (_, i) => String(r[i] ?? "").slice(0, 2000)));
    const name =
      text(a?.sheet_name)
        .replace(/[\\/?*[\]:]/g, " ")
        .trim()
        .slice(0, 28) || "Từ ứng dụng khác";
    return { op: "excel_new_sheet", name, values };
  }
  const title = text(a?.title).trim().slice(0, 300);
  const body = text(a?.body ?? a?.bullets ?? a?.text)
    .trim()
    .slice(0, 4000);
  if (!title && !body) throw new Error("Thiếu title hoặc body của slide cần thêm.");
  return { op: "ppt_add_slide", title, body };
}

/** Lệnh đọc từ kho có đúng hình dạng và đúng ứng dụng không (khung nhận kiểm lại, không tin khung gửi) */
export function validCommand(c: any, self: AppName, now = Date.now()): c is Command {
  if (!c || typeof c !== "object") return false;
  if (typeof c.id !== "string" || !/^[a-z0-9]{8,40}$/.test(c.id)) return false;
  if (!isApp(c.from) || c.to !== self || c.from === self) return false;
  if (typeof c.at !== "number" || now - c.at > CMD_MAX_AGE_MS || c.at - now > CMD_MAX_AGE_MS) return false;
  if (c.toTitle !== undefined && (typeof c.toTitle !== "string" || c.toTitle.length > 200)) return false;
  if (Object.keys(c).some((k) => !["id", "from", "to", "at", "toTitle", "payload"].includes(k))) return false;
  const p = c.payload;
  if (!p || p.op !== OP_OF[self]) return false;
  try {
    // Chuẩn hóa lại bằng chính bộ kiểm của bên gửi: sai hình dạng, quá cỡ đều bị loại
    const again = buildPayload(self, p.op === "excel_new_sheet" ? { values: p.values, sheet_name: p.name } : p);
    return JSON.stringify(again) === JSON.stringify(p);
  } catch {
    return false;
  }
}

/** Mô tả lệnh để hỏi người dùng ở khung gửi và để ghi nhận ở khung nhận */
/** `full` = bản cho hộp hỏi xác nhận: hiện gần đủ nội dung và nói rõ phần bị cắt (người dùng phải thấy thứ sắp được ghi,
 *  không chỉ câu mở đầu); bản ngắn dùng cho dòng ghi nhận ở khung nhận. */
export function describePayload(p: CrossOp, full = false): string {
  const cut = (s: string, n: number) => (s.length > n ? `${s.slice(0, n)}… [còn ${s.length - n} ký tự nữa chưa hiện ở đây]` : s);
  if (p.op === "word_append") {
    const n = p.text.split(/\n+/).filter((l) => l.trim()).length;
    return `thêm ${n} đoạn (${p.text.length} ký tự) vào cuối văn bản: “${cut(full ? p.text.replace(/\n+/g, " ¶ ") : p.text.replace(/\s+/g, " "), full ? 1500 : 160)}”`;
  }
  if (p.op === "excel_new_sheet") {
    const row = (r: string[]) => cut(r.join(" | "), full ? 200 : 120);
    const rows = p.values;
    const shown = full && rows.length > 1 ? `; ${rows.slice(1, 4).map((r, i) => `hàng ${i + 2}: ${row(r)}`).join("; ")}${rows.length > 5 ? `; … ; hàng cuối: ${row(rows[rows.length - 1])}` : rows.length === 5 ? `; hàng 5: ${row(rows[4])}` : ""}` : "";
    return `tạo sheet mới “${p.name}” với bảng ${rows.length}×${rows[0]?.length || 0} ô (hàng đầu: ${row(rows[0])}${shown})`;
  }
  return `thêm một slide: “${cut(p.title || "(không tiêu đề)", full ? 300 : 80)}”${p.body ? ` — ${cut(full ? p.body.replace(/\n+/g, " ¶ ") : p.body.replace(/\s+/g, " "), full ? 1200 : 120)}` : ""}`;
}

// ---------- Gửi ----------

export const newId = () => (Date.now().toString(36) + Math.random().toString(36).slice(2, 10)).toLowerCase().replace(/[^a-z0-9]/g, "");

export interface SendOptions {
  timeoutMs?: number;
  pollMs?: number;
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
  signal?: { aborted: boolean };
}

/** Lý do không gửi được ngay lúc này (khung đích chưa mở, không nhận ghi); null = gửi được */
export function blockReason(s: KVStore, from: AppName, to: AppName, now = Date.now()): string | null {
  if (from === to) return `Đang ở ${from}: dùng công cụ của chính ứng dụng này để ghi.`;
  const l = liveOf(s, to, now);
  if (!l) return `Chưa ghi được sang ${to}: khung AI Đa Nguồn ở ${to} chưa mở (hoặc ${to} chưa mở tệp). Nhờ người dùng mở ${to} cùng tệp cần ghi, bấm nút AI Đa Nguồn trên thẻ Home ở đó, rồi thử lại.`;
  if (!l.accept) return `Chưa ghi được sang ${to}: khung AI Đa Nguồn ở ${to} đang tắt “Nhận ghi từ ứng dụng khác”. Nhờ người dùng bật mục đó trong Cài đặt của khung ở ${to}, rồi thử lại.`;
  return null;
}

/** Đặt lệnh vào hộp thư của ứng dụng đích rồi chờ kết quả. Không ném lỗi vì hết hạn: trả ok=false kèm lý do rõ. */
export async function sendCommand(s: KVStore, from: AppName, to: AppName, payload: CrossOp, o: SendOptions = {}): Promise<CommandResult> {
  const now = o.now || Date.now;
  const sleep = o.sleep || ((ms: number) => new Promise<void>((r) => setTimeout(r, ms)));
  const block = blockReason(s, from, to, now());
  if (block) return { id: "", ok: false, at: now(), error: block };
  // Hộp thư chỉ chứa một lệnh: lệnh trước chưa được nhận thì không ghi đè
  const pending = read<Command>(s, CMD_PREFIX + to);
  if (pending && typeof pending.at === "number" && now() - pending.at <= CMD_MAX_AGE_MS) {
    return { id: "", ok: false, at: now(), error: `${to} đang có một lệnh ghi khác chờ xử lý. Thử lại sau vài giây.` };
  }
  const title = liveOf(s, to, now())?.title;
  const cmd: Command = { id: newId(), from, to, at: now(), ...(title ? { toTitle: title } : {}), payload };
  if (!write(s, CMD_PREFIX + to, cmd)) return { id: cmd.id, ok: false, at: now(), error: "Không ghi được lệnh vào bộ nhớ dùng chung của add-in (bộ nhớ đầy?)." };
  const deadline = now() + (o.timeoutMs ?? SEND_TIMEOUT_MS);
  const resKey = RES_PREFIX + cmd.id;
  try {
    while (now() < deadline) {
      const res = read<CommandResult>(s, resKey);
      if (res && res.id === cmd.id) return { id: cmd.id, ok: res.ok === true, at: res.at, message: res.message ? String(res.message).slice(0, 600) : undefined, error: res.error ? String(res.error).slice(0, 600) : undefined };
      if (o.signal?.aborted) break;
      await sleep(o.pollMs ?? 150);
    }
    // Hết hạn hoặc bị dừng: rút lệnh lại nếu khung đích chưa lấy, để nó không chạy muộn sau khi đã báo thất bại
    const still = read<Command>(s, CMD_PREFIX + to);
    const taken = !still || still.id !== cmd.id;
    if (!taken) s.removeItem(CMD_PREFIX + to);
    return {
      id: cmd.id,
      ok: false,
      at: now(),
      error: taken
        ? `${to} đã nhận lệnh nhưng chưa báo kết quả sau ${Math.round((o.timeoutMs ?? SEND_TIMEOUT_MS) / 1000)} giây. Có thể nội dung ĐÃ được ghi: nhờ người dùng xem tài liệu ở ${to} trước khi ghi lại, không gửi lại ngay.`
        : `${to} không nhận lệnh trong ${Math.round((o.timeoutMs ?? SEND_TIMEOUT_MS) / 1000)} giây (khung AI Đa Nguồn ở đó có thể vừa đóng hoặc đang bận). Lệnh đã được rút lại, chưa ghi gì.`,
    };
  } finally {
    try {
      s.removeItem(resKey);
    } catch {
      /* bỏ qua */
    }
  }
}

// ---------- Nhận ----------

/** Lấy lệnh đang chờ cho ứng dụng này (và gỡ khỏi hộp thư). Lệnh sai hình dạng, quá hạn, đã chạy đều bị bỏ.
 *  `seen` là tập id đã chạy của khung này. Trả null khi không có gì để làm. */
export function takeCommand(s: KVStore, self: AppName, o: { accept: boolean; seen: Set<string>; now?: number; title?: string }): Command | null {
  const now = o.now ?? Date.now();
  const raw = read<any>(s, CMD_PREFIX + self);
  if (!raw) return null;
  // Lệnh ghi cho một tệp khác đang mở ở cùng ứng dụng: để nguyên cho khung của tệp đó lấy
  if (raw && typeof raw.toTitle === "string" && raw.toTitle && o.title && raw.toTitle !== o.title) return null;
  const drop = () => {
    try {
      s.removeItem(CMD_PREFIX + self);
    } catch {
      /* bỏ qua */
    }
  };
  if (!validCommand(raw, self, now) || o.seen.has(raw.id)) {
    drop();
    // Lệnh bị bỏ vẫn báo lại (nếu có mã lệnh), để khung gửi không phải chờ hết hạn rồi đoán
    if (raw && typeof raw.id === "string" && /^[a-z0-9]{8,40}$/.test(raw.id) && !o.seen.has(raw.id))
      write(s, RES_PREFIX + raw.id, { id: raw.id, ok: false, at: now, error: `${self} đã bỏ lệnh ghi vì không hợp lệ hoặc quá hạn; chưa ghi gì.` } as CommandResult);
    return null;
  }
  drop();
  o.seen.add(raw.id);
  if (!o.accept) {
    write(s, RES_PREFIX + raw.id, { id: raw.id, ok: false, at: now, error: `Khung AI Đa Nguồn ở ${self} đang tắt “Nhận ghi từ ứng dụng khác”.` } as CommandResult);
    return null;
  }
  return raw as Command;
}

/** Hai khung cùng ứng dụng (hai tài liệu cùng tên hoặc chưa đặt tên) có thể cùng lấy một lệnh: mỗi khung ghi dấu nhận của mình,
 *  chờ một nhịp rồi đọc lại — chỉ khung còn thấy dấu của mình mới được chạy. */
export const CLAIM_PREFIX = "aidn.claim.";
export async function claimCommand(s: KVStore, id: string, token: string, wait: (ms: number) => Promise<void> = (ms) => new Promise((r) => setTimeout(r, ms))): Promise<boolean> {
  const key = CLAIM_PREFIX + id;
  try {
    const prior = s.getItem(key);
    if (prior && prior !== token) return false;
    s.setItem(key, token);
    await wait(80 + Math.floor(Math.random() * 60));
    const won = s.getItem(key) === token;
    // Dấu nhận tự hết tác dụng cùng lệnh; dọn sau khi lệnh chắc chắn đã quá hạn
    setTimeout(() => {
      try {
        if (s.getItem(key) === token) s.removeItem(key);
      } catch {
        /* bỏ qua */
      }
    }, CMD_MAX_AGE_MS + 5000);
    return won;
  } catch {
    return true;
  }
}

export function reply(s: KVStore, id: string, r: { ok: boolean; message?: string; error?: string }, now = Date.now()): void {
  write(s, RES_PREFIX + id, { id, at: now, ok: r.ok, message: r.message, error: r.error } as CommandResult);
}

/** Dọn kết quả cũ không ai nhận (khung gửi đã đóng) */
export function sweep(s: KVStore & { length?: number; key?: (i: number) => string | null }, now = Date.now()): number {
  let n = 0;
  try {
    const keys: string[] = [];
    for (let i = 0; i < (s.length || 0); i++) {
      const k = s.key?.(i);
      if (k && k.startsWith(RES_PREFIX)) keys.push(k);
    }
    for (const k of keys) {
      const v = read<CommandResult>(s, k);
      if (!v || typeof v.at !== "number" || now - v.at > 60000) {
        s.removeItem(k);
        n++;
      }
    }
  } catch {
    /* bỏ qua */
  }
  return n;
}
