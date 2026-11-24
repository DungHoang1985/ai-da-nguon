/** Hai chế độ dùng (Đợt 9, `tasks/dot/dot-9/dot-9.md`): "Miễn phí" chỉ hiện nguồn / model có bậc miễn phí và tự xoay vòng giữa các nguồn đó;
 *  "Trả phí" chỉ hiện nguồn / model thu phí, không xoay vòng. Hàm thuần, kiểm thử bằng npm test. */
import type { Provider } from "./providers";

export type UsageMode = "free" | "paid";

/** Thứ tự xoay vòng mặc định của chế độ Miễn phí; người dùng sửa được trong Cài đặt. OpenRouter đứng đầu từ 22/11/2026 (anh duyệt):
 *  so với add-in Claude chính thức, OpenRouter Nemotron 3 Ultra :free có Nghĩ kỹ đạt 97% điểm ở Excel, 94% ở PowerPoint; Gemini Flash
 *  ở Excel chỉ đạt 63% (so-sanh-openrouter/ket-qua/BAO-CAO-22-11.md). Trước đó (19/11): Gemini → Groq → OpenRouter → Cohere. */
export const DEFAULT_FREE_ORDER = ["openrouter", "gemini", "groq", "cohere"];

/** Gemini: chỉ dòng Flash, Flash-Lite và Gemma có bậc miễn phí; dòng Pro, ảnh, giọng nói thì không (R2 của dot-9). */
const GEMINI_FREE = /^(models\/)?(gemini-[\w.-]*flash|gemma)/i;
const GEMINI_NOT_FREE = /pro|image|tts|live|audio|embedding|veo|imagen/i;

/** Nguồn có bậc miễn phí. Không khai `tier` = chỉ trả phí. Nguồn tự khai: người dùng chọn khi khai. */
export function providerTier(p: Provider): "free" | "paid" | "both" {
  return p.tier || "paid";
}

export function providerInMode(p: Provider, mode: UsageMode): boolean {
  const t = providerTier(p);
  return t === "both" || t === mode;
}

/** Model có được hiện ở chế độ này không. Nguồn "both": Miễn phí chỉ nhận model có bậc miễn phí; Trả phí nhận mọi model,
 *  trừ OpenRouter (model `:free` thuộc chế độ Miễn phí). */
export function modelInMode(p: Provider, m: { id: string; free?: boolean }, mode: UsageMode): boolean {
  if (!providerInMode(p, mode)) return false;
  if (p.id === "openrouter") return mode === "free" ? !!m.free : !m.free;
  if (mode === "paid") return true;
  if (p.id === "gemini") return GEMINI_FREE.test(m.id) && !GEMINI_NOT_FREE.test(m.id);
  return true;
}

/** Dòng nhắc hiện khi nhập key ở chế độ Miễn phí: add-in không biết key thuộc gói nào (R1 của dot-9). */
export function freeKeyWarning(p: Provider): string {
  if (p.id === "openrouter") return "";
  const base = "Chỉ miễn phí khi key thuộc gói miễn phí (dự án / tài khoản chưa bật thanh toán). Đã bật thanh toán thì mỗi lần gọi bị tính tiền.";
  if (p.id === "cohere") return "Key dùng thử: miễn phí, giới hạn tốc độ, Cohere không cho dùng vào mục đích thương mại. " + base;
  return providerTier(p) === "paid" ? "" : base;
}

/** Thứ tự xoay vòng đã lưu, bỏ mã lạ / trùng, chỉ giữ nguồn có bậc miễn phí; thiếu nguồn nào thì nối vào cuối theo thứ tự mặc định. */
export function normalizeFreeOrder(saved: string[] | undefined, providers: Provider[]): string[] {
  const free = providers.filter((p) => providerInMode(p, "free")).map((p) => p.id);
  const out: string[] = [];
  for (const id of [...(saved || []), ...DEFAULT_FREE_ORDER, ...free]) if (free.includes(id) && !out.includes(id)) out.push(id);
  return out;
}
