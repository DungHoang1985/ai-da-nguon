/** Hạn mức và số dư của key OpenRouter — phần thuần, kiểm thử được bằng npm test.
 *
 *  Đối chiếu tài liệu 15/11/2026:
 *   - GET /api/v1/key (openrouter.ai/docs/api/reference/limits): trần của key, mức đã dùng, số lượt miễn phí trong ngày.
 *   - GET /api/v1/credits (…/api-reference/credits/get-credits): số dư tài khoản, nay đòi khóa quản lý; key thường nhận 403.
 *     Gọi thử một lần, không được thì thẻ chỉ hiện phần của /key và đường dẫn tới trang số dư. */
import type { Provider } from "./providers";
import { ensureOk, fetchDiagnosed } from "./errors";

export interface KeyInfo {
  label?: string;
  /** Trần chi tiêu đặt cho key (USD); null = không đặt trần */
  limit: number | null;
  /** Còn lại dưới trần của key; null = không đặt trần */
  limitRemaining: number | null;
  /** Đã dùng (USD): từ trước tới nay, hôm nay (giờ UTC), tháng này */
  usage?: number;
  usageDaily?: number;
  usageMonthly?: number;
  /** Tài khoản chưa từng nạp tiền */
  isFreeTier?: boolean;
  /** Số lượt gọi model miễn phí trong ngày (giờ UTC) */
  free?: { used?: number; limit?: number; remaining?: number };
  /** Số dư tài khoản (USD) — chỉ có khi key được xem /credits */
  balance?: number;
  /** Lúc đọc */
  at: number;
}

const num = (v: any): number | undefined => (typeof v === "number" && Number.isFinite(v) ? v : undefined);
const numOrNull = (v: any): number | null => num(v) ?? null;

/** Chuẩn hóa phản hồi GET /key */
export function parseKeyInfo(json: any, at = Date.now()): KeyInfo {
  const d = json?.data && typeof json.data === "object" ? json.data : json || {};
  const f = d.free_model_daily_requests;
  const free = f && typeof f === "object" ? { used: num(f.used), limit: num(f.limit), remaining: num(f.remaining) } : undefined;
  if (free && free.remaining === undefined && free.limit !== undefined && free.used !== undefined) free.remaining = Math.max(0, free.limit - free.used);
  return {
    label: typeof d.label === "string" ? d.label : undefined,
    limit: numOrNull(d.limit),
    limitRemaining: numOrNull(d.limit_remaining),
    usage: num(d.usage),
    usageDaily: num(d.usage_daily),
    usageMonthly: num(d.usage_monthly),
    isFreeTier: typeof d.is_free_tier === "boolean" ? d.is_free_tier : undefined,
    ...(free ? { free } : {}),
    at,
  };
}

/** Số dư từ phản hồi GET /credits: tổng đã nạp trừ tổng đã dùng */
export function parseBalance(json: any): number | undefined {
  const d = json?.data && typeof json.data === "object" ? json.data : json || {};
  const total = num(d.total_credits);
  const used = num(d.total_usage);
  return total === undefined || used === undefined ? undefined : total - used;
}

export function keyInfoRequest(p: Provider, key: string): { url: string; init: RequestInit } {
  return { url: `${p.baseUrl}/key`, init: { method: "GET", headers: { Authorization: `Bearer ${key}` } } };
}

export function creditsRequest(p: Provider, key: string): { url: string; init: RequestInit } {
  return { url: `${p.baseUrl}/credits`, init: { method: "GET", headers: { Authorization: `Bearer ${key}` } } };
}

/** Nguồn có xem được hạn mức bằng key không (hiện chỉ OpenRouter) */
export function hasKeyInfo(p?: Provider): boolean {
  return p?.id === "openrouter";
}

/** Đọc hạn mức của key. `tryBalance`: gọi thêm /credits; bị từ chối (key thường) thì bỏ qua, không coi là lỗi. */
export async function fetchKeyInfo(
  p: Provider,
  key: string,
  o: { fetchFn?: typeof fetch; tryBalance?: boolean; now?: number; signal?: AbortSignal } = {}
): Promise<{ info: KeyInfo; balanceAllowed: boolean | undefined }> {
  const fetchFn = o.fetchFn || fetch;
  const a = keyInfoRequest(p, key);
  const res = await fetchDiagnosed(p, a.url, { ...a.init, signal: o.signal }, fetchFn);
  await ensureOk(res, p);
  const info = parseKeyInfo(await res.json(), o.now ?? Date.now());
  // true = xem được; false = máy chủ nói rõ key này không được xem (401/403); undefined = chưa biết (không hỏi, mất mạng, máy chủ lỗi)
  let balanceAllowed: boolean | undefined;
  if (o.tryBalance) {
    try {
      const b = creditsRequest(p, key);
      const r = await fetchFn(b.url, { ...b.init, signal: o.signal });
      if (r.ok) {
        const bal = parseBalance(await r.json());
        if (bal !== undefined) {
          info.balance = bal;
          balanceAllowed = true;
        }
      } else if (r.status === 401 || r.status === 403) balanceAllowed = false;
    } catch {
      /* không xem được số dư lần này: thẻ vẫn hiện phần còn lại, lần sau thử lại */
    }
  }
  return { info, balanceAllowed };
}

/** Số tiền kiểu Việt: 1,23 USD; số rất nhỏ giữ 4 chữ số thập phân */
export function usd(n: number): string {
  const abs = Math.abs(n);
  const digits = abs !== 0 && abs < 0.01 ? 4 : 2;
  return `${n.toFixed(digits).replace(".", ",")} USD`;
}

export type CreditLevel = "ok" | "warn" | "out";

export interface CreditView {
  level: CreditLevel;
  /** Các dòng hiện trong thẻ hạn mức */
  lines: string[];
  /** Cảnh báo (mức warn / out), mỗi câu một ý */
  warnings: string[];
}

/** Ngưỡng cảnh báo */
export const LOW_FREE_REQUESTS = 5;
export const LOW_USD = 1;

/** Dựng nội dung thẻ hạn mức và các cảnh báo từ dữ liệu của key. */
export function describeKeyInfo(info: KeyInfo, o: { balanceAllowed?: boolean } = {}): CreditView {
  const lines: string[] = [];
  const warnings: string[] = [];
  let level: CreditLevel = "ok";
  const raise = (l: CreditLevel) => {
    if (l === "out" || (l === "warn" && level === "ok")) level = l;
  };

  if (info.balance !== undefined) {
    lines.push(`Số dư tài khoản: ${usd(info.balance)}`);
    if (info.balance < 0) {
      raise("out");
      warnings.push("Số dư tài khoản đang âm: OpenRouter từ chối mọi yêu cầu, kể cả model miễn phí, cho tới khi nạp thêm.");
    } else if (info.balance < LOW_USD && info.isFreeTier === false) {
      raise("warn");
      warnings.push(`Số dư tài khoản còn dưới ${usd(LOW_USD)}: model trả tiền sắp không gọi được.`);
    }
  } else if (o.balanceAllowed === false) {
    lines.push("Số dư tài khoản: key này không được xem (OpenRouter chỉ cho khóa quản lý xem). Bấm “Xem số dư” để mở trang của OpenRouter.");
  }

  if (info.limit !== null && info.limitRemaining === null) {
    // Có trần mà máy chủ không báo còn lại bao nhiêu: không đoán là đã hết
    lines.push(`Trần của key: ${usd(info.limit)}`);
  } else if (info.limit !== null) {
    const left = info.limitRemaining ?? 0;
    lines.push(`Trần của key: còn ${usd(left)} trên ${usd(info.limit)}`);
    if (left <= 0) {
      raise("out");
      warnings.push("Key đã chạm trần chi tiêu đặt cho nó: nâng trần ở trang quản lý key, hoặc đăng nhập lại để nhận key mới.");
    } else if (left < LOW_USD) {
      raise("warn");
      warnings.push(`Trần của key còn dưới ${usd(LOW_USD)}.`);
    }
  } else lines.push("Trần của key: không đặt trần");

  if (info.usageDaily !== undefined || info.usageMonthly !== undefined) {
    const parts: string[] = [];
    if (info.usageDaily !== undefined) parts.push(`hôm nay ${usd(info.usageDaily)}`);
    if (info.usageMonthly !== undefined) parts.push(`tháng này ${usd(info.usageMonthly)}`);
    lines.push(`Key này đã dùng: ${parts.join(" · ")}`);
  }

  const f = info.free;
  if (f && f.limit !== undefined) {
    const left = f.remaining ?? Math.max(0, f.limit - (f.used ?? 0));
    lines.push(`Model miễn phí hôm nay: còn ${left}/${f.limit} lượt (tính lại lúc 7 giờ sáng giờ Việt Nam)`);
    if (left <= 0) {
      raise(info.isFreeTier === false ? "warn" : "out");
      warnings.push("Đã hết lượt dùng model miễn phí của hôm nay. Có lại lúc 7 giờ sáng; chọn model trả tiền hoặc đổi nguồn.");
    } else if (left <= LOW_FREE_REQUESTS) {
      raise("warn");
      warnings.push(`Chỉ còn ${left} lượt dùng model miễn phí hôm nay.`);
    }
    if (f.limit <= 50 && info.isFreeTier !== false) lines.push("Tài khoản đã từng nạp từ 10 credit thì được 1.000 lượt miễn phí mỗi ngày.");
  }
  return { level, lines, warnings };
}

/** Một câu ngắn cho khung chat khi đang dùng OpenRouter và hạn mức thấp; không có gì đáng báo thì trả undefined.
 *  `freeModel`: model đang chọn là model miễn phí (chỉ khi đó số lượt miễn phí mới liên quan). */
export function chatCreditWarning(info: KeyInfo | undefined, o: { freeModel?: boolean; balanceAllowed?: boolean } = {}): { text: string; level: CreditLevel } | undefined {
  if (!info) return undefined;
  if (info.balance !== undefined && info.balance < 0) return { level: "out", text: "OpenRouter: số dư tài khoản đang âm, mọi yêu cầu sẽ bị từ chối cho tới khi nạp thêm." };
  const capLeft = info.limit !== null && info.limitRemaining !== null ? info.limitRemaining : undefined;
  if (capLeft !== undefined && capLeft <= 0) return { level: "out", text: "OpenRouter: key đã chạm trần chi tiêu. Nâng trần hoặc đăng nhập lại trong Cài đặt." };
  const f = info.free;
  if (o.freeModel && f && f.limit !== undefined) {
    const left = f.remaining ?? Math.max(0, f.limit - (f.used ?? 0));
    if (left <= 0) return { level: "out", text: "OpenRouter: đã hết lượt model miễn phí hôm nay (có lại lúc 7 giờ sáng). Chọn model trả tiền hoặc đổi nguồn." };
    if (left <= LOW_FREE_REQUESTS) return { level: "warn", text: `OpenRouter: còn ${left}/${f.limit} lượt model miễn phí hôm nay.` };
  }
  if (!o.freeModel) {
    if (capLeft !== undefined && capLeft < LOW_USD) return { level: "warn", text: `OpenRouter: trần của key còn ${usd(capLeft)}.` };
    if (info.balance !== undefined && info.balance < LOW_USD) return { level: "warn", text: `OpenRouter: số dư tài khoản còn ${usd(info.balance)}.` };
  }
  return undefined;
}
