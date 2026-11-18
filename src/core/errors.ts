/** Đổi lỗi kỹ thuật thành thông báo tiếng Việt đúng nguyên nhân. */
import type { Provider } from "./providers";

/** Lý do về key và tiền (Đợt 2): có lý do thì thông báo nêu đúng nguyên nhân và việc nên làm, và không tự thử lại. */
export type ErrorReason =
  /** Tài khoản hết tiền / hết tín dụng */
  | "no_credit"
  /** Key sai, nhà cung cấp không nhận */
  | "key_invalid"
  /** Key hết hạn, bị tắt hoặc bị thu hồi */
  | "key_expired"
  /** Chạm trần chi tiêu đặt cho key, workspace hoặc tổ chức */
  | "key_limit"
  /** OpenRouter: còn tiền nhưng không đủ cho độ dài trả lời tối đa của yêu cầu này */
  | "low_credit"
  /** OpenRouter: hết lượt dùng model miễn phí trong ngày */
  | "free_limit"
  /** Anthropic: key chưa gắn workspace */
  | "workspace"
  /** Anthropic: tổ chức không cho gọi trực tiếp từ trình duyệt */
  | "browser_blocked"
  /** Một yêu cầu lớn hơn hạn mức token mỗi phút của gói (Groq gói miễn phí trả 413): gửi lại bao nhiêu lần cũng thế */
  | "too_large";

export class ApiError extends Error {
  /** Lỗi tạm thời dù mã trạng thái không thuộc nhóm tự thử lại (OpenRouter: 402 do ngân sách của các yêu cầu đang chạy) */
  public transient?: boolean;
  /** Hết hạn mức theo NGÀY (429): gửi lại trong vài giây cũng vậy → không tự thử lại */
  public daily?: boolean;
  constructor(
    message: string,
    public status?: number,
    /** Số giây nhà cung cấp yêu cầu chờ trước khi gọi lại (tiêu đề retry-after), nếu đọc được */
    public retryAfter?: number,
    /** Lý do về key và tiền, nếu nhận ra */
    public reason?: ErrorReason
  ) {
    super(message);
  }
}

/** Lỗi tạm thời nên tự thử lại: hết hạn mức tạm thời (429), máy chủ lỗi hoặc quá tải (5xx, 529 của Anthropic).
 *  Lỗi đã rõ lý do về key hoặc tiền thì không thử lại: gửi lại bao nhiêu lần cũng thế (Gemini báo hết tiền bằng mã 429). */
export function isRetryable(e: unknown): boolean {
  // Lỗi mạng còn nguyên dạng TypeError đã qua fetchDiagnosed: máy chủ không với tới được lúc đó (mạng chập chờn) → thử lại.
  // Nhà cung cấp chặn gọi từ khung add-in (CORS) thì fetchDiagnosed đã đổi sang ApiError không thử lại.
  if (e instanceof TypeError) return FETCH_FAILED.test(String(e.message).trim());
  if (!(e instanceof ApiError)) return false;
  if (e.reason || e.daily) return false;
  if (e.transient) return true;
  const s = e.status;
  return typeof s === "number" && (s === 408 || s === 429 || s >= 500);
}

/** Câu lỗi 429 nói hết hạn mức theo ngày (Gemini "PerDay", OpenRouter "free-models-per-day", Groq "requests per day (RPD)") */
const DAILY_QUOTA = /per[\s_-]?day|PerDay|\bRPD\b|\bdaily\b/i;
const PER_MINUTE = /per[\s_-]?min|PerMinute|\bRPM\b|\bTPM\b/i;

/** Che mọi thứ trông như API key trong một câu trước khi hiện, chép hay lưu (soát mã độc lập 16/11/2026, T1):
 *  câu lỗi của nhà cung cấp có thể nhắc lại key, mà thông báo lỗi được hiện trên khung, vào nút "Chép chi tiết lỗi"
 *  và được lưu cùng hội thoại. Che cả dạng nhà cung cấp đã che dở (sk-proj-****abcd). */
/** Các key đang dùng trong phiên (chỉ trong bộ nhớ): che ĐÚNG giá trị đó ở mọi câu đưa ra ngoài, kể cả key không có tiền tố
 *  nhận ra được (JWT, chuỗi hex — MiniMax, Mistral, Cohere, nguồn tự khai) khi nhà cung cấp nhắc lại trong câu lỗi. */
const knownSecrets = new Set<string>();
export function rememberSecret(value: string | undefined | null): void {
  const v = String(value ?? "").trim();
  if (v.length >= 12) knownSecrets.add(v);
}
export function forgetSecrets(): void {
  knownSecrets.clear();
}

export function maskSecrets(text: string): string {
  let t = String(text ?? "");
  for (const k of knownSecrets) if (t.includes(k)) t = t.split(k).join("…[key]");
  return t
    .replace(/\bsk-[A-Za-z0-9_*.\-]{6,}/g, "sk-…")
    .replace(/\bAIza[A-Za-z0-9_\-]{16,}/g, "AIza…")
    .replace(/\b(gsk|xai|csk|nvapi|pplx|r8|hf)[_-][A-Za-z0-9_\-]{8,}/g, "$1_…")
    // Mã tổ chức / dự án của tài khoản (Groq "org_01…", OpenAI "org-…", "proj_…"): thông tin riêng, câu lỗi 413 và 429 hay nhắc lại
    .replace(/\b(org|proj)[_-](?=[A-Za-z0-9-]*\d)[A-Za-z0-9]{4,}(?:-[A-Za-z0-9]{4,})*(?![A-Za-z0-9_])/gi, (m, p) => (m.length - p.length - 1 >= 12 ? `${p}_…` : m))
    .replace(/\b(Bearer)\s+[A-Za-z0-9._~+/\-]{12,}=*/gi, "$1 …")
    .replace(/\b(api[ _-]?key|x-api-key|x-goog-api-key|access[ _-]?token|token)(['"]?\s*[:=]\s*['"]?)[A-Za-z0-9._~+/\-]{16,}/gi, "$1$2…");
}

/** Dấu hiệu chắc chắn của hết tiền: nhận ở mọi mã trạng thái */
const NO_CREDIT_HARD =
  /(credit balance is too low|prepayment credits are depleted|credits? (are|is|have been|has been) (depleted|exhausted|used up)|insufficient[ _-]?(balance|credits?|funds)|out of credits?|not enough credits?|negative (credit )?balance)/i;
/** Dấu hiệu mềm: chỉ nhận khi mã trạng thái không phải 429 (câu báo hết hạn mức hay kèm lời mời "purchase credits") */
const NO_CREDIT_SOFT = /(payment required|billing[ _](hard[ _])?limit|purchase (more )?credits|add more credits)/i;
/** OpenRouter: còn tiền nhưng không đủ cho độ dài trả lời tối đa của yêu cầu này */
const LOW_CREDIT = /(requires more credits|can only afford|fewer max_tokens)/i;
const FREE_DAILY = /free[- ]models?[- ]per[- ]day/i;
/** 401: mọi chữ này đều nghĩa là key không còn dùng được */
const KEY_GONE = /(expired|revoked|disabled|deactivated|inactive|been deleted|was deleted|archived|reported as leaked|has been blocked)/i;
/** 400 / 403: chỉ nhận khi chữ đó nói về chính key (đứng ngay sau "key" hoặc "token" trong cùng câu) */
const KEY_GONE_OF_KEY = /\b(api[ _-]?key|key|token)\b[^.]{0,40}\b(expired|revoked|disabled|deactivated|inactive|deleted|archived|leaked|blocked)\b/i;
const KEY_BAD =
  /(invalid (x-)?api[ _-]?key|incorrect api key|api key not valid|api[ _-]?key[^.]{0,40}(invalid|not found|malformed)|invalid (authentication|credentials|bearer|token)|no auth credentials|user not found|authentication (failed|error)|invalid_api_key|api_key_invalid)/i;
const SPEND_LIMIT = /(specified (api |workspace )?usage limits?|spend(ing)? limit|key limit (exceeded|reached)|budget (limit|exceeded|exhausted)|usage cap)/i;
/** Yêu cầu lớn hơn hạn mức mỗi phút của gói (đo 16/11/2026, Groq gói miễn phí trong Word: "Request too large for model … on tokens per minute (TPM): Limit 8000, Requested 8300") */
const TOO_LARGE = /request too large[\s\S]{0,400}?tokens per min/i;

/** Số trong câu "Limit 8000, Requested 8300" (nếu có) */
export function tpmNumbers(detail: string): { limit?: number; requested?: number } {
  const l = /\bLimit:?\s*([\d.,]+)/i.exec(detail);
  const r = /\bRequested:?\s*([\d.,]+)/i.exec(detail);
  const n = (m: RegExpExecArray | null) => (m ? Number(m[1].replace(/[.,](?=\d{3}\b)/g, "").replace(/[.,]$/, "")) || undefined : undefined);
  return { limit: n(l), requested: n(r) };
}
const RESTRICTED = /(only available|not available|not allowed|not permitted|agentic|your region|country|verif)/i;

/** Nhận ra lý do về key và tiền từ mã trạng thái, câu báo lỗi và thân phản hồi (nếu đọc được).
 *  Đối chiếu tài liệu chính thức 15/11/2026: Anthropic (errors, api-credits-for-subscribers, authentication),
 *  OpenRouter (errors-and-debugging, limits), và câu lỗi thật đã gặp của Gemini, OpenAI, DeepSeek. */
export function classifyError(status: number | undefined, detail = "", body?: any): { reason?: ErrorReason; transient?: boolean } {
  const err = body && typeof body === "object" ? (body.error && typeof body.error === "object" ? body.error : body) : {};
  const type = String(err.type || "");
  const code = String(err.code ?? "");
  const src = String(err.metadata?.limit_source || "");
  const reasons: string[] = Array.isArray(err.details) ? err.details.map((x: any) => String(x?.reason || "")) : [];

  // OpenRouter: nguồn của giới hạn nói rõ nhất
  if (src === "openrouter_in_flight_budget") return { transient: true };
  // "requires more credits, or fewer max_tokens … can only afford N": còn tiền nhưng thiếu cho yêu cầu này, không phải hết tiền
  if (LOW_CREDIT.test(detail)) return { reason: "low_credit" };
  if (src === "openrouter_credits") return { reason: "no_credit" };
  if (src === "openrouter_key_limit" || src === "openrouter_guardrail_budget" || src === "openrouter_workspace_budget") return { reason: "key_limit" };

  // Anthropic: hai câu có nghĩa riêng, đứng trước mọi phép đoán theo mã trạng thái
  if (/^\s*CORS requests are not allowed for this Organization/i.test(detail)) return { reason: "browser_blocked" };
  if (/not scoped to a workspace/i.test(detail)) return { reason: "workspace" };

  if (FREE_DAILY.test(detail)) return { reason: "free_limit" };
  // 413, hoặc 429 mà số token yêu cầu đã vượt trần mỗi phút: chờ cũng không qua
  if (TOO_LARGE.test(detail)) {
    const n = tpmNumbers(detail);
    if (status === 413 || (n.limit && n.requested && n.requested > n.limit)) return { reason: "too_large" };
  }
  // "exceeded your current quota": với OpenAI là hết tiền và có mã insufficient_quota (nhận ở dòng dưới); với Gemini là hết hạn mức
  // theo phút hoặc theo ngày, tự thử lại được — nên câu này không nằm trong các mẫu hết tiền.
  if (type === "billing_error" || code === "insufficient_quota" || type === "insufficient_quota") return { reason: "no_credit" };
  if (NO_CREDIT_HARD.test(detail)) return { reason: "no_credit" };
  if (status !== 429 && NO_CREDIT_SOFT.test(detail)) return { reason: "no_credit" };
  if (status === 402) return { reason: "no_credit" };
  if (SPEND_LIMIT.test(detail)) return { reason: "key_limit" };

  // 400 / 403: chỉ coi là lỗi về key khi câu nói rõ về key (có chữ "api key" ở đâu đó trong câu là chưa đủ — soát mã 16/11/2026, Th3)
  const keyGone = KEY_GONE_OF_KEY.test(detail);
  const apiKeyReason = reasons.some((r) => /API_KEY/.test(r));
  if ((status === 400 || status === 403) && keyGone) return { reason: "key_expired" };
  if (apiKeyReason || (status === 400 && KEY_BAD.test(detail) && /key/i.test(detail)) || KEY_BAD.test(code)) return { reason: "key_invalid" };
  if (status === 401) {
    // 401 còn được dùng khi model bị giới hạn nơi dùng (đo 15/11/2026, OpenRouter): để messageForStatus xử lý
    if (RESTRICTED.test(detail)) return {};
    return { reason: KEY_GONE.test(detail) ? "key_expired" : "key_invalid" };
  }
  return {};
}

/** Câu tiếng Việt cho từng lý do: nêu nguyên nhân, việc nên làm, và nhắc rằng hội thoại còn nguyên. */
export function messageForReason(reason: ErrorReason, p?: Provider, detail = ""): string {
  const name = p?.name || "Nhà cung cấp";
  const d = detail ? ` (${detail.slice(0, 200)})` : "";
  const keep = "Hội thoại vẫn còn nguyên";
  const swap = "đổi sang nguồn khác ở nhãn model cạnh ô nhập rồi gửi lại";
  const relogin = p?.oauth ? " hoặc bấm “Đăng nhập lại”" : "";
  switch (reason) {
    case "no_credit":
      return (
        `${name}: tài khoản đã hết tiền (hết tín dụng) nên yêu cầu bị từ chối. ${keep}: ${swap}, hoặc nạp thêm ở trang của ${name}.` +
        (p?.kind === "anthropic" && !p.dialect ? " Nếu anh đang dùng tín dụng API kèm gói Claude thì tín dụng kỳ này đã hết, kỳ sau có lại." : "") +
        d
      );
    case "low_credit":
      return (
        `${name}: số dư còn ít, không đủ cho độ dài trả lời tối đa của yêu cầu này (tài khoản chưa hết tiền). ` +
        `Nạp thêm ở trang của ${name}, chọn model rẻ hơn, hoặc ${swap}. ${keep}.${d}`
      );
    case "key_invalid":
      return `Key không đúng: ${name} không nhận key này. Vào Cài đặt, dán lại key${relogin}.${d}`;
    case "key_expired":
      return `Key của ${name} đã hết hạn, bị tắt hoặc bị thu hồi. Vào Cài đặt, dán key mới${relogin}; trong lúc đó có thể ${swap}.${d}`;
    case "key_limit":
      return `${name}: key đã chạm trần chi tiêu đặt cho key (hoặc cho workspace, tổ chức). Nâng trần ở trang của ${name}, chờ kỳ sau, hoặc ${swap}. ${keep}.${d}`;
    case "free_limit":
      return (
        `${name}: đã hết lượt dùng model miễn phí của hôm nay (50 lượt mỗi ngày; 1.000 lượt nếu tài khoản đã từng nạp từ 10 credit). ` +
        `Có lại lúc 7 giờ sáng giờ Việt Nam. Chọn model trả tiền hoặc ${swap}. ${keep}.${d}`
      );
    case "workspace":
      return (
        `Key của ${name} chưa gắn với workspace nào nên add-in không gọi được. Vào platform.claude.com → API keys, tạo key mới và chọn một workspace ` +
        `(Default là được), rồi dán vào Cài đặt.`
      );
    case "too_large": {
      const n = tpmNumbers(detail);
      const nf = (v: number) => Math.round(v).toLocaleString("vi-VN");
      const nums = n.limit && n.requested ? ` (yêu cầu này cần khoảng ${nf(n.requested)} token, gói chỉ cho ${nf(n.limit)} token mỗi phút)` : "";
      return (
        `${name}: yêu cầu lớn hơn hạn mức token mỗi phút của gói đang dùng${nums}. Mỗi yêu cầu của add-in gửi kèm mô tả các công cụ nên không thu nhỏ được; ` +
        `gửi lại cũng không qua. ${keep}: ${swap}, hoặc nâng gói ở trang của ${name}.`
      );
    }
    case "browser_blocked":
      return (
        `Tổ chức của key ${name} này không cho gọi trực tiếp từ add-in (do thiết lập lưu dữ liệu hoặc Compliance API của tổ chức); key không sai. ` +
        `Dùng key của tổ chức khác, hoặc dùng model này qua OpenRouter.`
      );
  }
}

/** Thông báo cho một lỗi của nhà cung cấp: có lý do về key hoặc tiền thì dùng câu của lý do, không thì theo mã trạng thái. */
export function explainApiError(
  status: number | undefined,
  detail = "",
  o: { provider?: Provider; body?: any } = {}
): { text: string; reason?: ErrorReason; transient?: boolean } {
  const c = classifyError(status, detail, o.body);
  // Phân loại trên câu gốc, nhưng mọi câu đưa ra ngoài đều đã che key
  detail = maskSecrets(detail);
  if (c.reason) return { text: messageForReason(c.reason, o.provider, detail), reason: c.reason };
  if (c.transient) {
    const d = detail ? ` (${detail.slice(0, 200)})` : "";
    return { text: `${o.provider?.name || "Nhà cung cấp"} đang giữ chỗ cho các yêu cầu khác của tài khoản; thử lại sau ít giây${d}`, transient: true };
  }
  return { text: status ? messageForStatus(status, detail) : detail };
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

/** Gợi ý tiếng Việt cho lỗi Office.js đã biết nguyên nhân (đo trong Office). Không nhận ra thì trả null.
 *  Word thu nhỏ cửa sổ: Office.js từ chối chèn nhận xét (AccessDenied) và ô điền mẫu (GeneralException) — đo 16/11/2026, tasks/dot-3.md. */
export function officeErrorHint(tool: string, e: any): string | null {
  const code = String(e?.code || "");
  const msg = String(e?.message || e || "");
  if (tool === "word_add_comment" && (code === "AccessDenied" || /AccessDenied|access.*denied/i.test(msg))) {
    return (
      "Word từ chối thêm nhận xét. Nguyên nhân thường gặp nhất: cửa sổ Word đang bị THU NHỎ — khi đó Word không cho add-in thêm nhận xét (sửa chữ vẫn được). " +
      "Chưa thêm được nhận xét nào. Nhắn người dùng mở lại (khôi phục) cửa sổ Word rồi bảo làm lại; nếu cửa sổ đang mở bình thường thì văn bản có thể đang bị hạn chế chỉnh sửa (thẻ Review → Restrict Editing)."
    );
  }
  return null;
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
  /** Lý do về key và tiền (để giao diện gợi ý đổi nguồn hoặc mở Cài đặt) */
  reason?: ErrorReason;
}

/** Lỗi fetch thật (mất mạng, bị chặn CORS, luồng đứt) theo cách các trình duyệt báo. */
const FETCH_FAILED = /^(Failed to fetch|NetworkError when attempting to fetch resource\.?|Load failed|fetch failed|network error|The network connection was lost\.?|The Internet connection appears to be offline\.?|net::ERR_[A-Z_]+)/i;
/** Lỗi dựng tiêu đề HTTP do key có ký tự lạ. */
const HEADER_FAILED = /(ByteString|is an invalid header value|Invalid value|Headers\.append|non ISO-8859-1|Invalid character in header)/i;

export function describeError(e: unknown, ctx: { provider?: Provider } = {}): ErrorInfo {
  const err = e as any;
  const name = ctx.provider?.name || "nhà cung cấp";
  // Lớp chặn cuối: câu nào ra tới đây cũng được che key trước khi hiện, chép hoặc lưu cùng hội thoại
  const raw = maskSecrets(err?.message ? String(err.message) : String(e));
  const detail = [
    `${err?.name || typeof e}: ${raw}`,
    err?.status ? `HTTP ${err.status}` : "",
    ctx.provider ? `Nguồn: ${ctx.provider.name} (${ctx.provider.baseUrl})` : "",
    typeof err?.stack === "string" ? err.stack.split("\n").slice(1, 4).join("\n") : "",
  ]
    .filter(Boolean)
    .join("\n");

  if (e instanceof ApiError) return { kind: "api", text: raw, detail, status: e.status, reason: e.reason };
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

export async function ensureOk(res: Response, provider?: Provider): Promise<void> {
  if (res.ok) return;
  let detail = "";
  let body: any;
  try {
    const text = await res.text();
    try {
      const parsed = JSON.parse(text);
      // Gemini có lúc trả mảng một phần tử
      const j = Array.isArray(parsed) ? parsed[0] : parsed;
      body = j;
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
  const x = explainApiError(res.status, detail, { provider, body });
  const err = new ApiError(x.text, res.status, retryAfter, x.reason);
  if (x.transient) err.transient = true;
  if (res.status === 429) {
    const all = `${detail} ${JSON.stringify(body ?? "")}`;
    // Câu nói rõ hạn mức theo phút thì không tính là theo ngày, kể cả khi kèm câu quảng cáo "… requests per day"
    if (DAILY_QUOTA.test(all) && !PER_MINUTE.test(all)) err.daily = true;
  }
  throw err;
}
