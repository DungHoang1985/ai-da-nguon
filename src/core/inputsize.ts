/** Đo token đầu vào của một lần gửi trước khi gửi (Đợt 12, phần A — `tasks/dot-12.md`).
 *  Tính ĐỦ: lời nhắc hệ thống, tin nhắn, lệnh gọi công cụ, kết quả công cụ và bộ công cụ gửi kèm (trước Đợt 12 vòng xoay bỏ qua bộ công cụ —
 *  Word Trả phí tới khoảng 30.000 ký tự). Ảnh và tệp gửi nguyên (PDF quét) không tính: mỗi nhà cung cấp đếm khác nhau. Hàm thuần.
 *  Hai hệ số, đo bằng Chạy thật 22/11/2026 với OpenRouter Nemotron 3 Ultra (`tasks/dot-12.md`, A1): lời nhắc, công cụ, JSON (kết quả công cụ,
 *  tham số lệnh gọi) khoảng 2,6 ký tự một token; chữ thường của người dùng và tệp đính kèm tiếng Việt khoảng 3,6. Một hệ số chung (3) lệch
 *  −15% ở lượt nhỏ và +21% ở lượt có hợp đồng 50 trang; hai hệ số còn lệch −7,6% · +1,4% · +6,7%.
 *  Họ Claude tách chữ tiếng Việt dày hơn nhiều (đo cùng ngày với Claude Haiku 5.5: 1,3–1,6 ký tự một token cho chữ, khoảng 2,1 cho lời nhắc
 *  và công cụ) nên có hệ số riêng; dùng hệ số chung cho Claude thì đếm thiếu 29–54%. Họ khác (Gemini, Groq, Cohere…) chưa đo riêng: dùng hệ số chung. */
import type { AgentMessage, ToolSpec } from "./agent/types";

/** Ký tự một token: lời nhắc hệ thống, bộ công cụ, JSON */
export const CHARS_PER_TOKEN_STRUCT = 2.6;
/** Ký tự một token: chữ của người dùng, câu trả lời, tệp đính kèm */
export const CHARS_PER_TOKEN_TEXT = 3.6;
export interface TokenRatio {
  struct: number;
  text: number;
}
const DEFAULT_RATIO: TokenRatio = { struct: CHARS_PER_TOKEN_STRUCT, text: CHARS_PER_TOKEN_TEXT };
const CLAUDE_RATIO: TokenRatio = { struct: 2.1, text: 1.5 };

/** Chữ tiếng Việt có khoảng 20–23% ký tự ngoài bảng mã ASCII (đo trên hợp đồng thử), tiếng Anh gần 0 */
const VI_SHARE = 0.2;
/** Ký tự một token của chữ tiếng Anh với họ Claude (soát mã 22/11: hệ số 1,5 đo trên tiếng Việt làm chặn nhầm tệp tiếng Anh dài) */
const CLAUDE_EN_TEXT = 3.8;

function countNonAscii(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) > 127) n++;
  return n;
}

/** Hệ số chữ theo tỷ lệ ký tự ngoài ASCII: họ Claude đi từ 3,8 (tiếng Anh) tới 1,5 (tiếng Việt); họ khác giữ hệ số đo được */
function textRatio(r: TokenRatio, share: number): number {
  if (r.text >= CHARS_PER_TOKEN_TEXT) return r.text;
  const vi = Math.min(1, Math.max(0, share / VI_SHARE));
  return CLAUDE_EN_TEXT - (CLAUDE_EN_TEXT - r.text) * vi;
}

/** Hệ số theo họ model (tên model; Claude qua Anthropic, OpenRouter "anthropic/…", Bedrock "…claude…" đều nhận ra) */
export function tokenRatio(model?: string): TokenRatio {
  return /claude|anthropic/i.test(String(model || "")) ? CLAUDE_RATIO : DEFAULT_RATIO;
}
/** Nguồn có sức chứa nhỏ hơn số đo / 0,9 thì coi là không chứa nổi (chừa chỗ cho phần trả lời và sai số ước lượng) */
export const FIT_SHARE = 0.9;
/** Từ mức này khung hiện số đo của lượt (lượt bình thường không hiện) */
export const SHOW_SIZE_FROM = 20000;

function contentChars(c: any): number {
  if (typeof c === "string") return c.length;
  if (Array.isArray(c)) return c.reduce((n, p) => n + (typeof p === "string" ? p.length : typeof p?.text === "string" ? p.text.length : 0), 0);
  return c ? String(c).length : 0;
}

/** Số token ước lượng của một lần gửi: tin nhắn (kể cả lời nhắc hệ thống) + lệnh gọi, kết quả công cụ + bộ công cụ */
export function estimateInputTokens(messages: AgentMessage[], tools: ToolSpec[] = [], model?: string): number {
  const r = tokenRatio(model);
  let struct = 0;
  let text = 0;
  let nonAscii = 0;
  for (const m of messages as any[]) {
    if (m.role === "system" || m.role === "tool") struct += contentChars(m.content);
    else {
      text += contentChars(m.content);
      if (typeof m.content === "string") nonAscii += countNonAscii(m.content);
    }
    for (const tc of m.toolCalls || []) struct += JSON.stringify(tc.args || {}).length + 40;
  }
  for (const t of tools) struct += JSON.stringify({ name: t.name, description: t.description, parameters: t.parameters }).length;
  return Math.ceil(struct / r.struct + text / textRatio(r, text ? nonAscii / text : 1));
}

/** Model chứa nổi số token này không (không biết sức chứa thì coi như chứa nổi) */
export function fitsContext(tokens: number, context?: number): boolean {
  return !context || tokens <= context * FIT_SHARE;
}

export interface ModelInfo {
  id: string;
  name?: string;
  context?: number;
}

/** Tối đa `max` model cùng nguồn chứa nổi lượt này, nhỏ nhất trước (đủ dùng, thường rẻ hơn) */
export function biggerModels(models: ModelInfo[], tokens: number, current: string, max = 3): ModelInfo[] {
  return models
    .filter((m) => m.id !== current && m.context && fitsContext(tokens, m.context))
    .sort((a, b) => (a.context || 0) - (b.context || 0))
    .slice(0, max);
}

const n = (v: number) => Math.round(v).toLocaleString("vi-VN");

/** Câu báo khi model đang chọn không chứa nổi lượt gửi (chế độ Trả phí: chỉ báo và gợi ý — anh chọn 22/11/2026, không tự đổi model) */
export function tooLargeMessage(o: { tokens: number; context: number; model: string; alternatives: ModelInfo[] }): string {
  const alt = o.alternatives.length
    ? ` Model cùng nguồn chứa được: ${o.alternatives.map((m) => `${m.name || m.id} (${n(m.context || 0)})`).join(", ")} — bấm “Đổi model” để chọn.`
    : " Nguồn này không có model nào đủ sức chứa: bớt tệp đính kèm, bắt đầu hội thoại mới, hoặc đổi sang nguồn khác.";
  return `Chưa gửi: lượt này khoảng ${n(o.tokens)} token (gồm lời nhắc, công cụ, lịch sử hội thoại và tệp đính kèm), lớn hơn sức chứa của model ${o.model} (${n(o.context)} token).${alt}`;
}

/** Dòng báo số đo khi lượt lớn (chỉ để hiện) */
export function sizeNote(tokens: number, context?: number): string {
  return `Lượt này khoảng ${n(tokens)} token đầu vào${context ? ` — model chứa ${n(context)} token` : ""}.`;
}
