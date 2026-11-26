/** Xoay vòng nguồn miễn phí (Đợt 9, `tasks/dot/dot-9/dot-9.md` mục 4 bước 4). Chỉ dùng ở chế độ Miễn phí; danh sách ứng viên do bên gọi
 *  dựng từ nguồn có bậc miễn phí và đã có key, nên ở đây không bao giờ có nguồn trả phí.
 *  - Nguồn im lặng, hết lượt theo phút, máy chủ lỗi / quá tải, mất mạng → chuyển nguồn kế tiếp NGAY (không thử lại cùng nguồn),
 *    nguồn đó nghỉ 60 giây (hoặc lâu hơn theo retry-after).
 *  - Hết lượt theo ngày / tháng → nghỉ đến mốc đặt lại của nguồn; lưu lại để mở khung lại vẫn nhớ.
 *  - Yêu cầu quá lớn (413, vượt sức chứa), lỗi key / tiền, model không gọi được công cụ, lỗi khác → bỏ nguồn đó trong lượt này.
 *  - Hết nguồn → báo từng nguồn vì sao, khi nào dùng lại; không bao giờ chuyển sang trả phí. */
import { estimateInputTokens } from "../inputsize";
import { ApiError, isRetryable, maskSecrets } from "../errors";
import type { Provider } from "../providers";
import { quirksOf } from "../providers";
import { CallModelOptions, callModelStep, isToolUnsupportedError } from "./llm";
import type { AgentMessage, StepResult } from "./types";
import { modelInMode, providerInMode } from "../modes";
import { pickModel, toolReliabilityNote } from "../catalog";

export interface Candidate {
  provider: Provider;
  key: string;
  model: string;
  vision?: boolean;
  maxOutput?: number;
  context?: number;
}

export type FailKind = "stall" | "rate" | "server" | "network" | "daily" | "monthly" | "too_large" | "key" | "tools" | "other";

export interface Cooldown {
  until: number;
  reason: string;
}

/** Nơi lưu thời điểm nghỉ của từng nguồn (localStorage qua Storage). Thiếu = chỉ nhớ trong lượt. */
export interface CooldownStore {
  get(): Record<string, Cooldown>;
  set(v: Record<string, Cooldown>): void;
}

export interface SwitchInfo {
  from: string;
  to: string;
  /** Lý do bằng tiếng Việt */
  reason: string;
  /** Đã có chữ của nguồn trước hiện ra trên khung (phần đó không vào lịch sử gửi model) */
  partial: boolean;
}

/** Nghỉ sau lỗi tạm thời (anh chốt 19/11/2026: 429 theo phút xoay ngay, nguồn nghỉ 60 giây) */
export const SHORT_COOLDOWN_MS = 60000;
/** Hạn chờ im lặng ở chế độ Miễn phí (anh chốt 19/11/2026); bật "Suy nghĩ kỹ" thì giữ 150 giây */
export const FREE_STALL_MS = 60000;
export const FREE_STALL_THINK_MS = 150000;
/** Ước lượng token = số ký tự / 3; nguồn có sức chứa nhỏ hơn 90% ước lượng thì bỏ qua, không gọi thử */
const CHARS_PER_TOKEN = 3;
const FIT_SHARE = 0.9;

const MONTHLY = /per[\s_-]?month|monthly|\bmonth\b/i;
const PER_MINUTE = /per[\s_-]?min|\/\s*min(ute)?\b|PerMinute|\bRPM\b|\bTPM\b/i;
const TOO_LARGE = /context[\s_-]?(length|window)|maximum context|too many tokens|request too large|reduce the length|prompt is too long|input.{0,20}too long|exceeds?.{0,30}(token|context)/i;

export function classifyFailure(e: any): FailKind {
  if (e instanceof ApiError && e.stall) return "stall";
  if (e instanceof ApiError && e.daily) return "daily";
  const msg = String(e?.message || "");
  const status = e?.status;
  // Câu nói rõ hạn mức theo phút thì không tính là theo tháng, kể cả khi kèm câu "… 1000 calls / month" của key dùng thử (soát mã Đợt 9, L1)
  if (status === 429 && MONTHLY.test(msg) && !PER_MINUTE.test(msg)) return "monthly";
  if (e?.reason === "too_large" || status === 413 || ((status === 400 || status === 422 || status === undefined) && TOO_LARGE.test(msg))) return "too_large";
  if (e?.reason) return "key";
  if (isToolUnsupportedError(e)) return "tools";
  if (e instanceof TypeError && isRetryable(e)) return "network";
  if (status === 429) return "rate";
  if (isRetryable(e)) return "server";
  return "other";
}

const REASON: Record<FailKind, string> = {
  stall: "không phản hồi",
  rate: "hết lượt tạm thời",
  server: "máy chủ lỗi hoặc quá tải",
  network: "mất kết nối",
  daily: "hết lượt hôm nay",
  monthly: "hết lượt tháng này",
  too_large: "yêu cầu quá lớn so với nguồn này",
  key: "lỗi key hoặc tài khoản",
  tools: "model không gọi được công cụ",
  other: "lỗi",
};

/** Mốc nửa đêm kế tiếp theo một múi giờ (Intl). Ngày đổi giờ mùa hè có thể lệch 1 giờ — chấp nhận. */
export function nextMidnightIn(timeZone: string, now: number): number {
  try {
    const parts = new Intl.DateTimeFormat("en-US", { timeZone, hour12: false, hour: "2-digit", minute: "2-digit", second: "2-digit" }).formatToParts(new Date(now));
    const get = (t: string) => Number(parts.find((p) => p.type === t)?.value || 0) % 24;
    const since = (get("hour") * 3600 + Number(parts.find((p) => p.type === "minute")?.value || 0) * 60 + Number(parts.find((p) => p.type === "second")?.value || 0)) * 1000;
    return now - since + 86400000;
  } catch {
    return now + 86400000;
  }
}

/** Mốc đặt lại hạn mức ngày theo nguồn. [Chưa kiểm trang chính thức 19/11/2026 — xem dot-9 R8] Gemini: nửa đêm giờ Thái Bình Dương;
 *  OpenRouter: nửa đêm UTC; nguồn khác: 24 giờ kể từ lúc lỗi. */
export function dailyResetAt(providerId: string, now: number): number {
  const tz = quirksOf({ id: providerId }).dailyResetTz;
  return tz ? nextMidnightIn(tz, now) : now + 86400000;
}

export function monthlyResetAt(now: number): number {
  const d = new Date(now);
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1);
}

/** Ước lượng token của yêu cầu (chữ, tham số lệnh gọi); ảnh và tệp không tính */
export function estimateTokens(messages: AgentMessage[]): number {
  let chars = 0;
  for (const m of messages as any[]) {
    chars += String(m.content || "").length;
    for (const tc of m.toolCalls || []) chars += JSON.stringify(tc.args || {}).length + 40;
  }
  return Math.ceil(chars / CHARS_PER_TOKEN);
}

const label = (c: Candidate) => `${c.provider.name} · ${c.model}`;
const clock = (t: number) => new Date(t).toLocaleString("vi-VN", { hour: "2-digit", minute: "2-digit", day: "2-digit", month: "2-digit" });

/** Thời gian chờ nhà cung cấp hẹn trong câu lỗi: Groq "Please try again in 14m22.5s", Gemini "Please retry in 37.4s" (mili giây) */
export function retryHintMs(msg: string): number {
  const m = /(?:try again|retry)\s+in\s+((?:\d+(?:\.\d+)?\s*[hms]\s*)+)/i.exec(msg || "");
  if (!m) return 0;
  let ms = 0;
  for (const x of m[1].matchAll(/(\d+(?:\.\d+)?)\s*([hms])/gi)) {
    const u = x[2].toLowerCase();
    ms += Number(x[1]) * (u === "h" ? 3600000 : u === "m" ? 60000 : 1000);
  }
  return Math.ceil(ms);
}

/** Khóa thời gian nghỉ: hạn mức miễn phí tính theo MODEL (Gemini Flash và Flash-Lite mỗi model một hạn mức), trừ hạn mức ngày
 *  của OpenRouter tính chung cả tài khoản cho mọi model :free (soát mã Đợt 9, M1) */
export function cooldownKey(c: { provider: { id: string }; model: string }, kind?: FailKind): string {
  // Giá trị khóa đã lưu trên máy người dùng: hạn mức chung tài khoản → khóa là đúng mã nguồn ("openrouter"), không đổi
  return quirksOf(c.provider).accountDailyQuota && kind === "daily" ? c.provider.id : `${c.provider.id}|${c.model}`;
}

/** Một vòng xoay cho MỘT lượt (bên gọi tạo mới mỗi lượt). Bước kế tiếp của lượt bắt đầu từ nguồn đang chạy được.
 *  Hai loại "không dùng": nghỉ có hạn (lỗi tạm thời, hết lượt ngày / tháng) — hết hạn là dùng lại, kể cả trong cùng lượt;
 *  bỏ cả lượt (lỗi key, yêu cầu quá lớn, không gọi được công cụ, lỗi khác). */
export class FreeRotation {
  private idx = 0;
  /** Nguồn bị bỏ cả lượt này: mã nguồn → lý do */
  private excluded = new Map<string, string>();
  /** Lý do gần nhất của từng nguồn, để báo khi hết nguồn */
  private notes = new Map<string, string>();
  /** Thời gian nghỉ trong bộ nhớ (dùng khi không có kho lưu, và để không mất khi kho lưu lỗi) */
  private mem: Record<string, Cooldown> = {};
  /** Trạng thái suy nghĩ riêng từng nguồn (nguồn này từ chối tham số suy nghĩ không kéo theo nguồn khác) */
  private states = new Map<string, { plain?: boolean }>();
  private now: () => number;

  constructor(
    private candidates: Candidate[],
    private o: { store?: CooldownStore; now?: () => number; stallMs?: number } = {}
  ) {
    this.now = o.now || Date.now;
    if (!candidates.length) throw new ApiError("Chế độ Miễn phí chưa có nguồn nào có key. Nhập key của Gemini, Groq, OpenRouter hoặc Cohere trong Cài đặt.");
  }

  /** Nguồn đang dùng (hoặc sẽ dùng ở bước kế tiếp) */
  current(): Candidate {
    return this.candidates[this.idx];
  }

  list(): Candidate[] {
    return this.candidates.slice();
  }

  private cooldowns(): Record<string, Cooldown> {
    let stored: Record<string, Cooldown> = {};
    try {
      stored = { ...(this.o.store?.get() || {}) };
    } catch {
      /* kho lưu lỗi: dùng bản trong bộ nhớ */
    }
    return { ...stored, ...this.mem };
  }

  private setCooldown(key: string, cd: Cooldown): void {
    this.mem[key] = cd;
    try {
      // Đọc lại ngay trước khi ghi để giảm chỗ hai khung (Word, Excel) ghi đè lên nhau
      const all = { ...(this.o.store?.get() || {}) };
      const t = this.now();
      for (const k of Object.keys(all)) if (!(all[k]?.until > t)) delete all[k];
      all[key] = cd;
      this.o.store?.set(all);
    } catch {
      /* không lưu được: vẫn nhớ trong lượt này */
    }
  }

  /** Lý do nguồn chưa dùng được lúc này, nếu có */
  private blocked(c: Candidate, need: number): string | undefined {
    const x = this.excluded.get(c.provider.id);
    if (x) return x;
    const all = this.cooldowns();
    const t = this.now();
    for (const k of [cooldownKey(c), quirksOf(c.provider).accountDailyQuota ? c.provider.id : ""]) {
      const cd = k ? all[k] : undefined;
      if (cd && cd.until > t) return `${cd.reason}, dùng lại lúc ${clock(cd.until)}`;
    }
    if (c.context && need > c.context * FIT_SHARE) return `yêu cầu khoảng ${need.toLocaleString("vi-VN")} token, lớn hơn sức chứa của model (${c.context.toLocaleString("vi-VN")})`;
    return undefined;
  }

  private exhausted(): ApiError {
    const lines = this.candidates.map((c) => `${c.provider.name}: ${this.notes.get(c.provider.id) || "không dùng được"}`);
    // Đợt 12: mọi nguồn đều bị bỏ vì lượt quá lớn → nói rõ cách làm nhỏ lượt lại
    const allTooBig = this.candidates.length > 0 && this.candidates.every((c) => /sức chứa|quá lớn/.test(this.notes.get(c.provider.id) || ""));
    const tip = allTooBig ? " Lượt này quá lớn cho mọi nguồn miễn phí: bớt tệp đính kèm, bắt đầu hội thoại mới, hoặc chuyển sang chế độ Trả phí với model sức chứa lớn." : "";
    return new ApiError(
      `Đã thử hết các nguồn miễn phí — ${lines.join("; ")}.${tip} Add-in không tự chuyển sang nguồn trả phí. Hội thoại vẫn còn nguyên: chờ nguồn dùng lại được, nhập thêm key miễn phí trong Cài đặt, hoặc tự chuyển sang chế độ Trả phí rồi gửi lại.`
    );
  }

  /** Gọi model một bước qua vòng xoay. `input` giữ nguyên mọi thứ trừ nguồn, key, model, nhận ảnh, trần đầu ra (lấy theo ứng viên). */
  async call(input: Omit<CallModelOptions, "provider" | "key" | "model"> & Partial<CallModelOptions>, onSwitch?: (i: SwitchInfo) => void): Promise<StepResult> {
    // Đợt 12: tính cả bộ công cụ gửi kèm (trước chỉ đếm tin nhắn, thiếu tới khoảng 10.000 token ở Word Trả phí)
    // Ước lượng theo họ model của từng nguồn (Claude đếm chữ tiếng Việt dày gấp đôi): đo lại cho mỗi ứng viên
    const needOf = (c: Candidate) => estimateInputTokens(input.messages, input.tools || [], c.model);
    const n = this.candidates.length;
    let shown = false;
    let lastLabel = "";
    let lastReason = "";
    // Mốc bắt đầu cố định cho cả vòng thử: this.idx đổi theo nguồn đang thử, không được dùng làm mốc (bỏ sót nguồn)
    const start = this.idx;
    for (let tried = 0; tried < n; tried++) {
      const i = (start + tried) % n;
      const c = this.candidates[i];
      const why = this.blocked(c, needOf(c));
      if (why) {
        this.notes.set(c.provider.id, why);
        // Nguồn đang nghỉ (lỗi ở lượt trước) bị bỏ qua cũng phải báo đổi nguồn: trước đây bỏ qua lặng lẽ, nguồn khác trả lời thay
        // mà khung vẫn ghi model đã chọn (thử 24/11/2026: Ling, Particle, Gemma trên OpenRouter → Cohere trả lời)
        if (!lastLabel) {
          lastLabel = label(c);
          lastReason = why;
        }
        continue;
      }
      if (lastLabel) onSwitch?.({ from: lastLabel, to: label(c), reason: lastReason, partial: shown });
      this.idx = i;
      shown = false;
      let st = this.states.get(c.provider.id);
      if (!st) this.states.set(c.provider.id, (st = {}));
      try {
        return await callModelStep({
          ...(input as CallModelOptions),
          provider: c.provider,
          key: c.key,
          model: c.model,
          vision: c.vision,
          maxOutput: c.maxOutput,
          state: st,
          retryDelaysMs: [],
          stallMs: input.stallMs ?? this.o.stallMs ?? FREE_STALL_MS,
          foreignToolsAsText: true,
          // Tìm web của OpenRouter (plugin "web") tính tiền theo kết quả kể cả với model :free → không bật ở chế độ Miễn phí (soát mã Đợt 9, H1)
          webSearch: quirksOf(c.provider).webPlugin ? false : input.webSearch,
          onText: (d) => {
            shown = true;
            input.onText(d);
          },
        });
      } catch (e: any) {
        if (input.signal?.aborted || e?.name === "AbortError") throw e;
        const kind = classifyFailure(e);
        const t = this.now();
        const reason = REASON[kind];
        const hint = retryHintMs(String(e?.message || ""));
        if (kind === "daily") this.setCooldown(cooldownKey(c, kind), { until: hint > 0 ? t + hint : dailyResetAt(c.provider.id, t), reason });
        else if (kind === "monthly") this.setCooldown(cooldownKey(c, kind), { until: monthlyResetAt(t), reason });
        else if (kind === "stall" || kind === "rate" || kind === "server" || kind === "network") {
          const ra = typeof e?.retryAfter === "number" ? e.retryAfter * 1000 : 0;
          this.setCooldown(cooldownKey(c, kind), { until: t + Math.max(SHORT_COOLDOWN_MS, ra, hint), reason });
        }
        // Lỗi không thuộc nhóm đã biết: nêu câu lỗi (đã che key) để người dùng biết nguồn hỏng vì sao
        const detailed = kind === "other" || kind === "key" || kind === "tools";
        const shownReason = detailed ? `${reason}: ${maskSecrets(String(e?.message || e)).slice(0, 200)}` : reason;
        if (detailed || kind === "too_large") this.excluded.set(c.provider.id, shownReason);
        this.notes.set(c.provider.id, shownReason);
        lastLabel = label(c);
        lastReason = shownReason;
      }
    }
    throw this.exhausted();
  }
}


/** Dựng danh sách ứng viên cho chế độ Miễn phí: nguồn đang chọn (với model đang chọn) đứng đầu, rồi các nguồn khác theo thứ tự
 *  xoay vòng. Chỉ nhận nguồn có bậc miễn phí, đã có key, không bị chặn gọi từ add-in, và model thuộc chế độ Miễn phí.
 *  Model của nguồn khác: model đã nhớ nếu hợp lệ, không thì chọn từ danh mục đã lưu; chưa có danh mục thì bỏ nguồn đó. */
export function buildFreeCandidates(o: {
  order: string[];
  selected?: { providerId: string; modelId: string };
  keys: Record<string, string>;
  getProvider: (id: string) => Provider | undefined;
  modelCache: Record<string, { models: { id: string; free?: boolean; tools?: boolean; vision?: boolean; context?: number; maxOutput?: number; name: string; created?: number }[] } | undefined>;
  remembered: (pid: string) => string;
  agentOn: boolean;
}): Candidate[] {
  const ids = [...(o.selected ? [o.selected.providerId] : []), ...o.order];
  const out: Candidate[] = [];
  for (const pid of ids) {
    if (out.some((c) => c.provider.id === pid)) continue;
    const p = o.getProvider(pid);
    const key = o.keys[pid];
    if (!p || !key || p.blockedReason || !providerInMode(p, "free")) continue;
    const all = (o.modelCache[pid]?.models || []).filter((m) => modelInMode(p, m, "free"));
    const usable = o.agentOn ? all.filter((m) => m.tools !== false) : all;
    let id = o.selected?.providerId === pid ? o.selected.modelId : "";
    if (!usable.some((m) => m.id === id)) {
      // Model đã nhớ mà đã thấy không gọi công cụ đáng tin (⚠) thì không dùng cho nguồn dự phòng: nguồn dự phòng phải làm được việc
      const mem = o.remembered(pid);
      const remembered = o.agentOn && toolReliabilityNote(pid, mem) ? "" : mem;
      id = pickModel(usable, { providerId: pid, remembered, agentOn: o.agentOn, freeOnly: true });
    }
    const m = usable.find((x) => x.id === id);
    if (!m) continue;
    out.push({ provider: p, key, model: m.id, vision: m.vision ?? p.vision, maxOutput: m.maxOutput, context: m.context });
  }
  return out;
}
