/** Danh sách nhà cung cấp. Thêm nguồn OpenAI-compatible mới: chỉ cần thêm 1 dòng. */

export type ProviderKind = "openai" | "gemini" | "anthropic";

export interface Provider {
  id: string;
  name: string;
  kind: ProviderKind;
  /** URL gốc API (không có dấu / cuối) */
  baseUrl: string;
  /** Trang lấy API key */
  keyUrl: string;
  /** URL liệt kê model nếu khác `${baseUrl}/models` */
  modelsUrl?: string;
  /** Ghi chú hiển thị trong Cài đặt */
  note?: string;
  /** Nhà cung cấp không cho gọi trực tiếp từ add-in (đo được): hiện lý do ngay khi chọn */
  blockedReason?: string;
  /** Mọi model của nguồn này nhận ảnh (true) / không nhận (false); không ghi = tùy model hoặc chưa rõ */
  vision?: boolean;
}

/** Đo ngày 14 và 15/11/2026: máy chủ không trả tiêu đề Access-Control-Allow-Origin nên trình duyệt trong Office chặn. */
const NO_CORS = "nhà cung cấp này chặn gọi trực tiếp từ add-in (đo 15/11/2026: máy chủ không trả tiêu đề CORS). Dùng model này qua OpenRouter nếu cần.";

export const PROVIDERS: Provider[] = [
  {
    id: "openrouter",
    name: "OpenRouter",
    kind: "openai",
    baseUrl: "https://openrouter.ai/api/v1",
    keyUrl: "https://openrouter.ai/keys",
    note: "Nhiều model miễn phí (đuôi :free)",
  },
  {
    id: "gemini",
    name: "Google Gemini",
    kind: "gemini",
    baseUrl: "https://generativelanguage.googleapis.com/v1beta",
    keyUrl: "https://aistudio.google.com/apikey",
    vision: true,
  },
  {
    id: "groq",
    name: "Groq",
    kind: "openai",
    baseUrl: "https://api.groq.com/openai/v1",
    keyUrl: "https://console.groq.com/keys",
  },
  {
    id: "anthropic",
    name: "Anthropic (Claude)",
    kind: "anthropic",
    baseUrl: "https://api.anthropic.com/v1",
    keyUrl: "https://console.anthropic.com/settings/keys",
    vision: true,
  },
  {
    id: "openai",
    name: "OpenAI",
    kind: "openai",
    baseUrl: "https://api.openai.com/v1",
    keyUrl: "https://platform.openai.com/api-keys",
  },
  {
    id: "mistral",
    name: "Mistral",
    kind: "openai",
    baseUrl: "https://api.mistral.ai/v1",
    keyUrl: "https://console.mistral.ai/api-keys",
  },
  {
    id: "cerebras",
    name: "Cerebras",
    kind: "openai",
    baseUrl: "https://api.cerebras.ai/v1",
    keyUrl: "https://cloud.cerebras.ai/",
  },
  {
    id: "deepseek",
    name: "DeepSeek",
    kind: "openai",
    baseUrl: "https://api.deepseek.com/v1",
    keyUrl: "https://platform.deepseek.com/api_keys",
    vision: false,
  },
  {
    id: "sambanova",
    name: "SambaNova",
    kind: "openai",
    baseUrl: "https://api.sambanova.ai/v1",
    keyUrl: "https://cloud.sambanova.ai/apis",
    blockedReason: NO_CORS,
  },
  {
    id: "nvidia",
    name: "NVIDIA NIM",
    kind: "openai",
    baseUrl: "https://integrate.api.nvidia.com/v1",
    keyUrl: "https://build.nvidia.com/",
    blockedReason: NO_CORS,
  },
  {
    id: "cohere",
    name: "Cohere",
    kind: "openai",
    baseUrl: "https://api.cohere.ai/compatibility/v1",
    modelsUrl: "https://api.cohere.com/v1/models?endpoint=chat&page_size=1000",
    keyUrl: "https://dashboard.cohere.com/api-keys",
  },
];

export function getProvider(id: string): Provider | undefined {
  return PROVIDERS.find((p) => p.id === id);
}
