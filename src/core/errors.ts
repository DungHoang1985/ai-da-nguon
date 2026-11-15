/** Đổi lỗi kỹ thuật thành thông báo tiếng Việt đúng nguyên nhân. */
import type { Provider } from "./providers";

export class ApiError extends Error {
  constructor(
    message: string,
    public status?: number,
    /** Số giây nhà cung cấp yêu cầu chờ trước khi gọi lại (tiêu đề retry-after), nếu đọc được */
    public retryAfter?: number
  ) {
    super(message);
  }
}

/** Lỗi tạm thời nên tự thử lại: hết hạn mức tạm thời (429), máy chủ lỗi hoặc quá tải (5xx, 529 của Anthropic). */
export function isRetryable(e: unknown): boolean {
  const s = (e as any)?.status;
  return e instanceof ApiError && typeof s === "number" && (s === 429 || s >= 500);
}

export function messageForStatus(status: number, detail = ""): string {
  const d = detail ? ` (${detail.slice(0, 200)})` : "";
  if (status === 401 || status === 403) {
    // Không phải lúc nào cũng là sai key: nhà cung cấp còn dùng 401/403 khi model bị giới hạn nơi dùng, vùng, hoặc tài khoản chưa được cấp model
    // (đo 15/11/2026, OpenRouter: "… is only available on agentic harnesses").
    const restricted = /(only available|not available|not allowed|not permitted|agentic|your region|country|verif)/i.test(detail);
    const aboutKey = /(api[ _-]?key|token|credential|unauthori[sz]ed|authenticat)/i.test(detail);
    if (restricted || (status === 403 && detail && !aboutKey)) {
      return `Nhà cung cấp từ chối yêu cầu (${status}): tài khoản hoặc key không được dùng model này, hoặc model bị giới hạn nơi dùng. Hãy chọn model khác${d}`;
    }
    return status === 401 ? `Key không đúng hoặc đã hết hạn${d}` : `Key không đúng hoặc không có quyền${d}`;
  }
  if (status === 402) return `Tài khoản hết tiền hoặc cần nạp thêm${d}`;
  if (status === 404) return `Không tìm thấy model hoặc địa chỉ API${d}`;
  if (status === 429) return `Hết hạn mức (gọi quá nhiều). Thử lại sau hoặc đổi model${d}`;
  if (status === 529) return `Nhà cung cấp đang quá tải (529). Thử lại sau hoặc đổi model${d}`;
  if (status >= 500) return `Máy chủ nhà cung cấp đang lỗi (${status}). Thử lại sau${d}`;
  return `Lỗi ${status}${d}`;
}

/** Ký tự ẩn hay lẫn vào khi chép key từ web, Zalo, Word: dấu cách rộng 0, BOM, gạch nối mềm… */
const INVISIBLE = /[​-‍⁠﻿­]/g;

/** Làm sạch key trước khi lưu và trước khi dùng: bỏ ký tự ẩn, khoảng trắng hai đầu, dấu nháy bao ngoài. */
export function cleanKey(raw: string): string {
  let k = String(raw ?? "")
    .replace(INVISIBLE, "")
    .replace(/^[\s ]+|[\s ]+$/g, "");
  const q = /^(["'`“”‘’])(.*)\1$/.exec(k) || /^[“‘](.*)[”’]$/.exec(k);
  if (q) k = (q[2] ?? q[1]).trim();
  return k;
}

/** Key còn ký tự không gửi được trong tiêu đề HTTP (chữ có dấu, dấu cách, xuống dòng, "…") → lời nhắc; hợp lệ → null. */
export function keyProblem(key: string): string | null {
  if (!key) return null;
  if (/[^\x21-\x7E]/.test(key)) {
    return "Key có ký tự không hợp lệ (dấu cách, xuống dòng, chữ có dấu hoặc ký tự đặc biệt). Hãy sao chép lại key từ trang của nhà cung cấp rồi dán lại.";
  }
  return null;
}

export type ErrorKind = "aborted" | "api" | "key" | "network" | "app";

export interface ErrorInfo {
  kind: ErrorKind;
  /** Câu hiện cho người dùng */
  text: string;
  /** Chi tiết kỹ thuật để chép gửi khi cần tra lỗi */
  detail: string;
  status?: number;
}

/** Lỗi fetch thật (mất mạng, bị chặn CORS, luồng đứt) theo cách các trình duyệt báo. */
const FETCH_FAILED = /^(Failed to fetch|NetworkError when attempting to fetch resource\.?|Load failed|fetch failed|network error|The network connection was lost\.?|The Internet connection appears to be offline\.?|net::ERR_[A-Z_]+)/i;
/** Lỗi dựng tiêu đề HTTP do key có ký tự lạ. */
const HEADER_FAILED = /(ByteString|is an invalid header value|Invalid value|Headers\.append|non ISO-8859-1|Invalid character in header)/i;

export function describeError(e: unknown, ctx: { provider?: Provider } = {}): ErrorInfo {
  const err = e as any;
  const name = ctx.provider?.name || "nhà cung cấp";
  const raw = err?.message ? String(err.message) : String(e);
  const detail = [
    `${err?.name || typeof e}: ${raw}`,
    err?.status ? `HTTP ${err.status}` : "",
    ctx.provider ? `Nguồn: ${ctx.provider.name} (${ctx.provider.baseUrl})` : "",
    typeof err?.stack === "string" ? err.stack.split("\n").slice(1, 4).join("\n") : "",
  ]
    .filter(Boolean)
    .join("\n");

  if (e instanceof ApiError) return { kind: "api", text: e.message, detail, status: e.status };
  if (err?.name === "AbortError") return { kind: "aborted", text: "Đã dừng.", detail };
  if (e instanceof TypeError && HEADER_FAILED.test(raw)) {
    return {
      kind: "key",
      text: `Key của ${name} có ký tự lạ (dấu cách ẩn, chữ có dấu, xuống dòng…) nên không gửi được. Vào Cài đặt, dán lại key.`,
      detail,
    };
  }
  if (e instanceof TypeError && FETCH_FAILED.test(raw.trim())) {
    const text = ctx.provider?.blockedReason
      ? `${name}: ${ctx.provider.blockedReason}`
      : /network error/i.test(raw)
        ? `Mất kết nối với ${name} giữa chừng. Thử gửi lại.`
        : `Không kết nối được tới ${name}: mất mạng, hoặc nhà cung cấp chặn gọi trực tiếp từ add-in (CORS). Vào Cài đặt → Kiểm tra nhà cung cấp để biết chính xác.`;
    return { kind: "network", text, detail };
  }
  return { kind: "app", text: `Lỗi trong add-in: ${raw}`, detail };
}

/** Giữ tên cũ cho các chỗ đang dùng. */
export function networkErrorMessage(e: unknown, ctx: { provider?: Provider } = {}): string {
  return describeError(e, ctx).text;
}

/** Máy chủ trả lỗi nhưng trang lỗi thiếu tiêu đề CORS nên trình duyệt không cho đọc (đo 15/11/2026: OpenAI với key dạng
 *  "sk-…" sai/hết hạn trả 401 không kèm Access-Control-Allow-Origin; Cerebras trả 401 không kèm tiêu đề). */
export function hiddenErrorMessage(p: Provider, model?: string): string {
  const m = model ? `model “${model}”` : "model này";
  return (
    `${p.name} trả về lỗi nhưng không cho add-in đọc chi tiết. Thường do: key sai hoặc hết hạn, tài khoản hết tiền, ` +
    `hoặc tài khoản chưa được dùng ${m} (OpenAI yêu cầu xác minh tổ chức với một số model mới). Kiểm tra key và số dư trên trang ${p.name}, hoặc chọn model khác.`
  );
}

/** fetch có chẩn đoán: lỗi fetch mà máy chủ vẫn tới được (gọi thử chế độ no-cors) → máy chủ trả lỗi giấu chi tiết, không phải mất mạng. */
export async function fetchDiagnosed(p: Provider, url: string, init: RequestInit, fetchFn: typeof fetch = fetch, model?: string): Promise<Response> {
  try {
    return await fetchFn(url, init);
  } catch (e: any) {
    const aborted = (init.signal as AbortSignal | undefined)?.aborted || e?.name === "AbortError";
    if (aborted || !(e instanceof TypeError) || !FETCH_FAILED.test(String(e.message).trim()) || p.blockedReason) throw e;
    let reachable = false;
    try {
      await fetchFn(new URL(url).origin + "/", { method: "GET", mode: "no-cors", cache: "no-store" });
      reachable = true;
    } catch {
      reachable = false;
    }
    if (reachable) throw new ApiError(hiddenErrorMessage(p, model));
    throw e;
  }
}

export async function ensureOk(res: Response): Promise<void> {
  if (res.ok) return;
  let detail = "";
  try {
    const text = await res.text();
    try {
      const j = JSON.parse(text);
      detail = j?.error?.message || j?.message || j?.error || text;
      if (typeof detail !== "string") detail = JSON.stringify(detail);
    } catch {
      detail = text;
    }
  } catch {
    /* bỏ qua */
  }
  let retryAfter: number | undefined;
  try {
    const ra = Number(res.headers?.get("retry-after"));
    if (Number.isFinite(ra) && ra >= 0) retryAfter = ra;
  } catch {
    /* tiêu đề không đọc được (CORS) */
  }
  throw new ApiError(messageForStatus(res.status, detail), res.status, retryAfter);
}
