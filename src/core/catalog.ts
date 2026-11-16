/** Trạng thái danh mục model của nguồn đang chọn — hàm thuần, tách khỏi Chat.tsx để kiểm thử.
 *  Sửa lỗi đổi nhà cung cấp (kế hoạch v4 mục 3.3): đổi nguồn là xóa ngay danh sách và model cũ; bỏ kết quả tải về muộn;
 *  không cho gửi khi model không thuộc nguồn đang chọn; nhớ model theo từng nguồn. */
import type { ModelInfo } from "./models";
import type { Provider } from "./providers";

export type CatalogStatus = "loading" | "ready" | "error";

export interface CatalogState {
  providerId: string;
  status: CatalogStatus;
  /** Chỉ chứa model của providerId */
  models: ModelInfo[];
  /** Thời điểm danh sách được tải (ms) */
  at?: number;
  /** Lý do tải lỗi (khi có) */
  error?: string;
  /** Số hiệu lần tải: kết quả về muộn của lần cũ bị bỏ */
  token: number;
}

export function startCatalog(providerId: string, cached: { at: number; models: ModelInfo[] } | undefined, token: number): CatalogState {
  return cached?.models?.length
    ? { providerId, status: "ready", models: cached.models, at: cached.at, token }
    : { providerId, status: "loading", models: [], token };
}

export function applyLoaded(s: CatalogState, providerId: string, token: number, models: ModelInfo[], at: number): CatalogState {
  if (providerId !== s.providerId || token !== s.token) return s;
  return { providerId, status: "ready", models, at, token };
}

/** Tải lỗi: còn danh sách cũ (bộ nhớ đệm) thì vẫn dùng được, ghi kèm lý do; không còn gì thì là lỗi. */
export function applyFailed(s: CatalogState, providerId: string, token: number, error: string): CatalogState {
  if (providerId !== s.providerId || token !== s.token) return s;
  return { ...s, status: s.models.length ? "ready" : "error", error };
}

export interface FilterOptions {
  providerId: string;
  freeOnly: boolean;
  toolsOnly: boolean;
  agentOn: boolean;
  search: string;
}

/** "Miễn phí" và "Chỉ model thao tác được" chỉ có nghĩa với OpenRouter (nơi danh mục có dữ liệu đó). */
export function filterModels(models: ModelInfo[], o: FilterOptions): ModelInfo[] {
  const or = o.providerId === "openrouter";
  const q = o.search.trim().toLowerCase();
  return models
    .filter((m) => !(or && o.freeOnly) || m.free)
    .filter((m) => !(or && o.agentOn && o.toolsOnly) || m.tools !== false)
    .filter((m) => !q || m.name.toLowerCase().includes(q) || m.id.toLowerCase().includes(q));
}

/** Model nên chọn cho nguồn chưa từng dùng: theo thứ tự ưu tiên, bản mới nhất trước. Lối cuối: model gọi được công cụ. */
const PREFER: Record<string, RegExp[]> = {
  // Bốn mẫu đầu là model trả tiền; các mẫu sau để khi bật "Miễn phí" vẫn rơi vào dòng model quen, gọi được công cụ,
  // thay vì model miễn phí mới nhất bất kỳ (15/11/2026: rơi vào model chỉ cho dùng trong "agentic harnesses").
  openrouter: [
    /^anthropic\/claude-sonnet/i,
    /^openai\/gpt-5/i,
    /^google\/gemini-[\d.]+-flash/i,
    /^google\/gemini-[\d.]+-pro/i,
    /^openai\/gpt-oss-120b/i,
    /^meta-llama\/llama-3\.3-70b/i,
    /^qwen\/qwen3/i,
    /^deepseek\/deepseek-(chat|v3)/i,
  ],
  anthropic: [/claude-sonnet/i, /claude-opus/i, /claude-haiku/i],
  openai: [/^gpt-5(\.\d+)?$/i, /^gpt-5(\.\d+)?-mini$/i, /^gpt-4\.1$/i, /^gpt-4o$/i],
  gemini: [/^gemini-[\d.]+-flash$/i, /^gemini-[\d.]+-pro$/i, /^gemini-.*flash(?!.*(lite|image|tts|live))/i],
  groq: [/gpt-oss-120b/i, /llama-3\.3-70b/i, /kimi/i, /qwen/i],
  cerebras: [/gpt-oss-120b/i, /qwen-3/i, /llama-3\.3-70b/i],
  deepseek: [/^deepseek-chat$/i],
  mistral: [/^mistral-large-latest$/i, /^mistral-medium-latest$/i, /^mistral-small-latest$/i],
  cohere: [/^command-a-\d{2}-\d{4}$/i, /^command-a(?!.*(vision|translate|reasoning))/i, /^command-r-plus/i],
  sambanova: [/llama-3\.3-70b/i, /deepseek-v3/i],
  nvidia: [/llama-3\.3-70b/i, /nemotron/i],
};

/** OpenRouter, khi bật "Miễn phí": dòng model miễn phí đã Chạy thật gọi công cụ đúng trong Word, xếp trước mọi mẫu khác
 *  (anh duyệt 16/11/2026: Nemotron 3 Ultra free đúng 6/6 lần thử ở Đợt 2 và Đợt 3). */
const PREFER_FREE: RegExp[] = [/^nvidia\/nemotron-3-ultra/i];

/** Model đã thấy trả lời "đã sửa" mà KHÔNG gọi công cụ khi Chạy thật trong Word (danh mục vẫn ghi là gọi được công cụ).
 *  Chỉ thêm vào đây khi có bằng chứng Chạy thật; ghi ngày và nơi lưu bằng chứng. */
export const UNRELIABLE_TOOL_MODELS: { re: RegExp; note: string }[] = [
  { re: /^baidu\/ernie-5\.5-flash/i, note: "Đã thấy model này trả lời “Đã sửa” mà không gọi công cụ, tài liệu không đổi (thử 16/11/2026)." },
];

/** Lời cảnh báo cho một model đã thấy không gọi công cụ đáng tin; null = chưa thấy vấn đề */
export function toolReliabilityNote(providerId: string, modelId: string): string | null {
  if (providerId !== "openrouter") return null;
  return UNRELIABLE_TOOL_MODELS.find((x) => x.re.test(modelId))?.note || null;
}

function newestFirst(a: ModelInfo, b: ModelInfo): number {
  if ((b.created || 0) !== (a.created || 0)) return (b.created || 0) - (a.created || 0);
  return b.id.localeCompare(a.id);
}

export function pickModel(
  models: ModelInfo[],
  o: { providerId: string; remembered?: string; agentOn: boolean; freeOnly: boolean }
): string {
  if (!models.length) return "";
  if (o.remembered && models.some((m) => m.id === o.remembered)) return o.remembered;
  let pool = models;
  if (o.providerId === "openrouter" && o.freeOnly) pool = pool.filter((m) => m.free);
  if (o.agentOn) {
    const usable = pool.filter((m) => m.tools !== false);
    if (usable.length) pool = usable;
  }
  if (!pool.length) pool = models;
  if (o.agentOn && o.providerId === "openrouter") {
    // Đang cần AI thao tác: tránh model đã thấy không gọi công cụ, trừ khi không còn model nào khác
    const reliable = pool.filter((m) => !toolReliabilityNote(o.providerId, m.id));
    if (reliable.length) pool = reliable;
  }
  const prefer = [...(o.providerId === "openrouter" && o.freeOnly ? PREFER_FREE : []), ...(PREFER[o.providerId] || [])];
  for (const re of prefer) {
    const hit = pool.filter((m) => re.test(m.id)).sort(newestFirst);
    if (hit.length) return hit[0].id;
  }
  // Không khớp mẫu nào: model gọi được công cụ, bản mới nhất trước (không lấy mục đầu theo chữ cái)
  const withTools = pool.filter((m) => m.tools === true).sort(newestFirst);
  return (withTools[0] || [...pool].sort(newestFirst)[0]).id;
}

/** Lý do chưa gửi được (hiện ngay cạnh nút Gửi); null = gửi được. */
export function sendBlockReason(o: { provider?: Provider; catalog: CatalogState; modelId: string }): string | null {
  if (!o.provider) return "Chưa chọn nhà cung cấp.";
  const c = o.catalog;
  if (c.providerId !== o.provider.id) return "Đang đổi nhà cung cấp…";
  if (c.status === "loading") return `Đang tải danh sách model của ${o.provider.name}…`;
  if (c.status === "error") {
    return `${o.provider.name}: ${c.error || "không tải được danh sách model"}${o.provider.blockedReason ? "" : ". Bấm “Làm mới” để thử lại."}`;
  }
  if (!o.modelId || !c.models.some((m) => m.id === o.modelId)) return `Hãy chọn model của ${o.provider.name}.`;
  return null;
}
