/** Liệt kê model của từng nhà cung cấp và chuẩn hóa về một định dạng. */
import { Provider } from "./providers";
import { ensureOk } from "./errors";

export interface ModelInfo {
  id: string;
  name: string;
  free?: boolean;
  context?: number;
  /** Gọi được công cụ (thao tác trên tài liệu). undefined = không rõ */
  tools?: boolean;
}

const NON_CHAT = /(embed|whisper|tts|dall-e|moderation|rerank|transcribe|audio|image-gen|imagen|veo|text-to-speech|guard)/i;

function sortModels(list: ModelInfo[]): ModelInfo[] {
  const seen = new Set<string>();
  return list
    .filter((m) => m.id && !seen.has(m.id) && (seen.add(m.id), true))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** Chuẩn hóa phản hồi /models — tách riêng để kiểm thử với dữ liệu mẫu. */
export function normalizeModels(provider: Provider, json: any): ModelInfo[] {
  if (provider.kind === "gemini") {
    const list: any[] = json?.models || [];
    return sortModels(
      list
        .filter((m) => (m.supportedGenerationMethods || []).includes("generateContent"))
        .map((m) => {
          const id = String(m.name || "").replace(/^models\//, "");
          return { id, name: m.displayName ? `${m.displayName} (${id})` : id, context: m.inputTokenLimit };
        })
        .filter((m) => !NON_CHAT.test(m.id))
    );
  }
  if (provider.kind === "anthropic") {
    const list: any[] = json?.data || [];
    return sortModels(list.map((m) => ({ id: m.id, name: m.display_name ? `${m.display_name} (${m.id})` : m.id })));
  }
  // Cohere dùng API riêng: { models: [{ name, context_length }] }
  if (Array.isArray(json?.models) && !Array.isArray(json?.data)) {
    return sortModels(
      json.models.map((m: any) => ({ id: m.name || m.id, name: m.name || m.id, context: m.context_length }))
    );
  }
  const list: any[] = json?.data || [];
  return sortModels(
    list
      .filter((m) => !NON_CHAT.test(String(m.id)))
      .filter((m) => m.active !== false)
      .map((m) => {
        const id = String(m.id);
        const p = m.pricing;
        const free =
          provider.id === "openrouter"
            ? id.endsWith(":free") || (p && Number(p.prompt) === 0 && Number(p.completion) === 0)
            : undefined;
        const sp: any[] | undefined = Array.isArray(m.supported_parameters) ? m.supported_parameters : undefined;
        return {
          id,
          name: m.name && m.name !== id ? `${m.name} (${id})` : id,
          free: free || undefined,
          context: m.context_length || m.context_window || undefined,
          tools: sp ? sp.includes("tools") : undefined,
        };
      })
  );
}

export function modelsRequest(provider: Provider, key: string): { url: string; init: RequestInit } {
  if (provider.kind === "gemini") {
    return {
      url: `${provider.baseUrl}/models?pageSize=1000&key=${encodeURIComponent(key)}`,
      init: { method: "GET" },
    };
  }
  if (provider.kind === "anthropic") {
    return {
      url: `${provider.baseUrl}/models?limit=1000`,
      init: { method: "GET", headers: anthropicHeaders(key) },
    };
  }
  return {
    url: provider.modelsUrl || `${provider.baseUrl}/models`,
    init: { method: "GET", headers: { Authorization: `Bearer ${key}` } },
  };
}

export function anthropicHeaders(key: string): Record<string, string> {
  return {
    "x-api-key": key,
    "anthropic-version": "2023-06-01",
    "anthropic-dangerous-direct-browser-access": "true",
    "content-type": "application/json",
  };
}

export async function fetchModels(provider: Provider, key: string, fetchFn: typeof fetch = fetch): Promise<ModelInfo[]> {
  const { url, init } = modelsRequest(provider, key);
  const res = await fetchFn(url, init);
  await ensureOk(res);
  const json = await res.json();
  const models = normalizeModels(provider, json);
  if (!models.length) throw new Error("Nhà cung cấp không trả về model chat nào.");
  return models;
}

/** Danh mục quá 24 giờ thì làm mới. */
export const CATALOG_TTL_MS = 24 * 60 * 60 * 1000;
export function isStale(at: number | undefined, now = Date.now()): boolean {
  return !at || now - at > CATALOG_TTL_MS;
}
