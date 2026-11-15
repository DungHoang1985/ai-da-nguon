/** Đăng nhập để nhận API key (OAuth PKCE) — khối dùng chung, không phụ thuộc Office, kiểm thử được bằng npm test.
 *
 *  Luồng (đối chiếu openrouter.ai/docs/guides/overview/auth/oauth ngày 15/11/2026):
 *   1. Khung add-in tạo một mã kiểm bí mật (code_verifier) và mã thách thức = SHA-256 của nó (code_challenge).
 *   2. Người dùng mở trang đăng nhập của nhà cung cấp kèm mã thách thức, đăng nhập và bấm cho phép.
 *      - Lối hộp thoại: trang của nhà cung cấp chuyển về trang của add-in kèm ?code=…; trang đó gửi mã về khung.
 *      - Lối dán mã: không gửi callback_url, trang của nhà cung cấp hiện mã để chép rồi dán vào add-in.
 *   3. Khung đổi mã + mã kiểm lấy API key (POST), lưu key như key dán tay.
 *
 *  Quy tắc an toàn: mã kiểm chỉ nằm trong bộ nhớ của khung (không ghi kho lưu, không vào địa chỉ trang); thông báo lỗi
 *  không bao giờ chứa mã đăng nhập, mã kiểm hay key; trang hộp thoại chỉ chuyển tới đúng tên miền khai trong cấu hình. */
import type { OAuthConfig, Provider } from "./providers";

export const CHALLENGE_METHOD = "S256";
/** Mã đăng nhập của nhà cung cấp hết hạn sau 10 phút (tài liệu OpenRouter) */
export const CODE_TTL_MS = 10 * 60 * 1000;
/** Loại tin nhắn trang hộp thoại gửi về khung */
export const DIALOG_MESSAGE_TYPE = "aidn-oauth";

/** base64url không đệm (RFC 7636, phụ lục A) */
export function base64Url(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  const b64 = typeof btoa === "function" ? btoa(bin) : Buffer.from(bin, "binary").toString("base64");
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

type RandomFn = (n: number) => Uint8Array;

const secureRandom: RandomFn = (n) => {
  const c = (globalThis as any).crypto;
  if (!c?.getRandomValues) throw new OAuthError("Trình duyệt của Office không có bộ sinh số ngẫu nhiên an toàn nên không đăng nhập được. Hãy dán key theo cách cũ.");
  return c.getRandomValues(new Uint8Array(n));
};

/** Mã kiểm bí mật: 32 byte ngẫu nhiên → 43 ký tự (RFC 7636 yêu cầu 43–128 ký tự) */
export function newVerifier(random: RandomFn = secureRandom): string {
  return base64Url(random(32));
}

/** Chuỗi chống giả mạo yêu cầu (state): gửi đi và phải nhận lại đúng nó */
export function newState(random: RandomFn = secureRandom): string {
  return base64Url(random(16));
}

/** Mã thách thức S256 = base64url(SHA-256(mã kiểm)) */
export async function s256Challenge(verifier: string): Promise<string> {
  const subtle = (globalThis as any).crypto?.subtle;
  if (!subtle?.digest) throw new OAuthError("Trình duyệt của Office không tính được SHA-256 nên không đăng nhập được. Hãy dán key theo cách cũ.");
  const data = new TextEncoder().encode(verifier);
  return base64Url(new Uint8Array(await subtle.digest("SHA-256", data)));
}

/** Một lần đăng nhập đang mở. `verifier` là bí mật: chỉ giữ trong bộ nhớ. */
export interface LoginSession {
  verifier: string;
  challenge: string;
  state: string;
  createdAt: number;
}

export async function createLoginSession(now = Date.now(), random: RandomFn = secureRandom): Promise<LoginSession> {
  const verifier = newVerifier(random);
  return { verifier, challenge: await s256Challenge(verifier), state: newState(random), createdAt: now };
}

const B64URL = /^[A-Za-z0-9_-]+$/;
export const validChallenge = (s: string): boolean => B64URL.test(s) && s.length === 43;
export const validState = (s: string): boolean => B64URL.test(s) && s.length >= 8 && s.length <= 128;

/** Địa chỉ trang đăng nhập của nhà cung cấp. Có `callbackUrl` = lối hộp thoại; không có = nhà cung cấp hiện mã để chép. */
export function buildAuthUrl(cfg: OAuthConfig, o: { challenge: string; state?: string; callbackUrl?: string }): string {
  if (!validChallenge(o.challenge)) throw new OAuthError("Mã thách thức không hợp lệ.");
  const u = new URL(cfg.authUrl);
  if (u.protocol !== "https:" || u.hostname !== cfg.host) throw new OAuthError("Cấu hình đăng nhập không hợp lệ.");
  if (o.callbackUrl) u.searchParams.set("callback_url", o.callbackUrl);
  u.searchParams.set("code_challenge", o.challenge);
  u.searchParams.set("code_challenge_method", CHALLENGE_METHOD);
  if (cfg.keyLabel) u.searchParams.set("key_label", cfg.keyLabel);
  // state chỉ được trả về ở lối có callback_url
  if (o.callbackUrl && o.state) u.searchParams.set("state", o.state);
  return u.toString();
}

/** Địa chỉ mở trong hộp thoại Office: trang của chính add-in (bắt buộc cùng tên miền với khung), trang đó tự chuyển tới nhà cung cấp.
 *  Chỉ mang mã thách thức và state — không có bí mật nào trong địa chỉ. */
export function dialogStartUrl(pageUrl: string, providerId: string, s: Pick<LoginSession, "challenge" | "state">): string {
  const u = new URL(pageUrl);
  u.search = "";
  u.hash = "";
  u.searchParams.set("p", providerId);
  u.searchParams.set("c", s.challenge);
  u.searchParams.set("s", s.state);
  return u.toString();
}

export type DialogQuery =
  | { kind: "start"; providerId: string; challenge: string; state: string }
  | { kind: "callback"; code: string; state?: string }
  | { kind: "denied"; detail: string }
  | { kind: "none" };

/** Trang hộp thoại đọc địa chỉ của chính nó: lần mở đầu (p, c, s) hay lần nhà cung cấp chuyển về (code). */
export function parseDialogQuery(search: string): DialogQuery {
  const q = new URLSearchParams(search.replace(/^\?/, ""));
  const code = q.get("code");
  if (code) {
    const c = cleanCode(code);
    if (!c) return { kind: "none" };
    const st = q.get("state") || "";
    return { kind: "callback", code: c, ...(validState(st) ? { state: st } : {}) };
  }
  const err = q.get("error") || q.get("error_description");
  if (err) return { kind: "denied", detail: String(err).replace(/[^\w .,:;()\-]/g, " ").slice(0, 120) };
  const p = q.get("p") || "";
  const c = q.get("c") || "";
  const s = q.get("s") || "";
  if (/^[a-z0-9-]{2,40}$/.test(p) && validChallenge(c) && validState(s)) return { kind: "start", providerId: p, challenge: c, state: s };
  return { kind: "none" };
}

const INVISIBLE = /[​-‍⁠﻿­]/g;

/** Mã hợp lệ: chữ ASCII in được, không có khoảng trắng, 6–2048 ký tự. Sai thì trả chuỗi rỗng. */
function cleanCode(raw: string): string {
  let c = String(raw ?? "").replace(INVISIBLE, "").trim();
  const q = /^(["'`“”‘’])(.*)\1$/.exec(c) || /^[“‘](.*)[”’]$/.exec(c);
  if (q) c = (q[2] ?? q[1]).trim();
  return /^[\x21-\x7E]{6,2048}$/.test(c) ? c : "";
}

export type PastedCode = { kind: "code"; code: string } | { kind: "key"; key: string } | { kind: "invalid"; reason: string };

/** Chuỗi người dùng dán vào ô "Dán mã": nhận mã trần, cả địa chỉ có ?code=…; nhận ra khi người dùng dán nhầm API key. */
export function extractCode(pasted: string, p?: Provider): PastedCode {
  const raw = String(pasted ?? "").replace(INVISIBLE, "").trim();
  if (!raw) return { kind: "invalid", reason: "Chưa có mã. Bấm “Mở trang lấy mã”, chép mã trên trang rồi dán vào đây." };
  // Dán nguyên địa chỉ hoặc nguyên dòng "code=…"
  const m = /[?&#]code=([^&#\s]+)/.exec(raw) || /^code=([^&#\s]+)$/.exec(raw);
  let code = "";
  if (m) {
    try {
      code = cleanCode(decodeURIComponent(m[1]));
    } catch {
      code = cleanCode(m[1]);
    }
  } else if (/^https?:\/\//i.test(raw)) {
    return { kind: "invalid", reason: "Địa chỉ vừa dán không chứa mã. Chép riêng mã hiện trên trang rồi dán lại." };
  } else code = cleanCode(raw);
  if (!code) return { kind: "invalid", reason: "Mã không hợp lệ (có dấu cách, xuống dòng hoặc ký tự lạ). Chép lại mã trên trang rồi dán lại." };
  if (looksLikeApiKey(code, p)) return { kind: "key", key: code };
  return { kind: "code", code };
}

/** Người dùng dán API key vào ô mã (key OpenRouter bắt đầu bằng sk-or-) */
export function looksLikeApiKey(s: string, p?: Provider): boolean {
  if (/^sk-or-/i.test(s)) return true;
  return !p && /^sk-[A-Za-z0-9_-]{20,}$/.test(s);
}

/** Lỗi của luồng đăng nhập. `message` là câu tiếng Việt cho người dùng, không bao giờ chứa mã, mã kiểm hay key. */
export class OAuthError extends Error {
  constructor(
    message: string,
    public status?: number,
    /** Nên gợi ý lối dán mã (hộp thoại không dùng được) */
    public suggestPaste?: boolean
  ) {
    super(message);
    this.name = "OAuthError";
  }
}

/** Bỏ mọi bí mật khỏi một câu trước khi đưa vào thông báo */
export function redact(text: string, secrets: (string | undefined)[]): string {
  let out = String(text ?? "");
  for (const s of secrets) if (s && s.length >= 4) out = out.split(s).join("…");
  // Phòng khi máy chủ nhắc lại một key
  return out.replace(/sk-[A-Za-z0-9_-]{8,}/g, "sk-…");
}

export function exchangeRequest(cfg: OAuthConfig, code: string, verifier: string): { url: string; init: RequestInit } {
  return {
    url: cfg.exchangeUrl,
    init: {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ code, code_verifier: verifier, code_challenge_method: CHALLENGE_METHOD }),
    },
  };
}

/** Câu tiếng Việt cho lỗi đổi mã (mã lỗi theo tài liệu OpenRouter: 400, 403, 405). `detail` đã được bỏ bí mật. */
export function exchangeErrorMessage(name: string, status: number, detail = ""): string {
  const d = detail ? ` (${detail.slice(0, 160)})` : "";
  if (status === 403 && /expired/i.test(detail)) return `Mã đăng nhập ${name} đã hết hạn (mã chỉ dùng được trong 10 phút). Bấm đăng nhập lại để lấy mã mới.`;
  if (status === 403 || status === 401)
    return `${name} không nhận mã này: mã sai, đã dùng rồi, hoặc không phải mã lấy từ lần mở trang vừa rồi của khung này. Bấm đăng nhập lại để lấy mã mới${d}.`;
  if (status === 400) return `${name} từ chối yêu cầu đổi mã (400). Bấm đăng nhập lại; nếu vẫn lỗi thì dán key theo cách cũ${d}.`;
  if (status === 405) return `${name} từ chối kiểu gọi đổi mã (405). Đây là lỗi của add-in: dán key theo cách cũ và báo lại lỗi này.`;
  if (status === 429) return `${name} báo gọi quá nhiều (429). Chờ một phút rồi bấm đăng nhập lại.`;
  if (status >= 500) return `Máy chủ ${name} đang lỗi (${status}). Thử đăng nhập lại sau ít phút.`;
  return `Không đổi được mã đăng nhập ${name} (lỗi ${status})${d}.`;
}

/** Đổi mã lấy API key. Ném OAuthError với câu tiếng Việt; không ném, không ghi mã hay mã kiểm ra bất cứ đâu. */
export async function exchangeCode(p: Provider, code: string, verifier: string, fetchFn: typeof fetch = fetch): Promise<string> {
  if (!p.oauth) throw new OAuthError(`${p.name} không hỗ trợ đăng nhập.`);
  const { url, init } = exchangeRequest(p.oauth, code, verifier);
  let res: Response;
  try {
    res = await fetchFn(url, init);
  } catch {
    throw new OAuthError(`Không kết nối được tới ${p.name} để đổi mã: mất mạng hoặc bị chặn. Thử lại; nếu vẫn lỗi thì dán key theo cách cũ.`);
  }
  let text = "";
  try {
    text = await res.text();
  } catch {
    /* không đọc được thân */
  }
  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    json = undefined;
  }
  if (!res.ok) {
    const raw = json?.error?.message || json?.message || (typeof json?.error === "string" ? json.error : "") || "";
    throw new OAuthError(exchangeErrorMessage(p.name, res.status, redact(String(raw), [code, verifier])), res.status);
  }
  const key = typeof json?.key === "string" ? json.key.trim() : "";
  if (!key) throw new OAuthError(`${p.name} trả lời nhưng không kèm key. Bấm đăng nhập lại; nếu vẫn lỗi thì dán key theo cách cũ.`);
  return key;
}

/** Tin nhắn trang hộp thoại gửi về khung (chuỗi JSON) */
export function dialogMessage(q: { code: string; state?: string }): string {
  return JSON.stringify({ t: DIALOG_MESSAGE_TYPE, code: q.code, state: q.state || "" });
}

export type DialogResult = { ok: true; code: string } | { ok: false; reason: string };

/** Khung đọc tin nhắn của hộp thoại: đúng loại, đúng state của lần đăng nhập đang mở, mã hợp lệ. */
export function readDialogMessage(message: string, session: Pick<LoginSession, "state"> | null): DialogResult {
  let j: any;
  try {
    j = JSON.parse(message);
  } catch {
    return { ok: false, reason: "Cửa sổ đăng nhập gửi về dữ liệu không đọc được." };
  }
  if (!j || j.t !== DIALOG_MESSAGE_TYPE) return { ok: false, reason: "Cửa sổ đăng nhập gửi về dữ liệu lạ." };
  if (j.denied) return { ok: false, reason: "Chưa được cho phép ở trang của nhà cung cấp." };
  if (!session) return { ok: false, reason: "Lần đăng nhập này đã kết thúc. Bấm đăng nhập lại." };
  if (!j.state || j.state !== session.state) return { ok: false, reason: "Mã trả về không thuộc lần đăng nhập đang mở (state không khớp). Bấm đăng nhập lại." };
  const code = cleanCode(String(j.code || ""));
  if (!code) return { ok: false, reason: "Cửa sổ đăng nhập không trả về mã hợp lệ. Bấm đăng nhập lại." };
  return { ok: true, code };
}
