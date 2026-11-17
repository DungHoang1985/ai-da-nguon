/** Liệt kê model của từng nhà cung cấp và chuẩn hóa về một định dạng. */
import { Provider } from "./providers";
import { ensureOk, fetchDiagnosed } from "./errors";

export interface ModelInfo {
  id: string;
  name: string;
  free?: boolean;
  context?: number;
  /** Gọi được công cụ (thao tác trên tài liệu). undefined = không rõ */
  tools?: boolean;
  /** Nhận ảnh. undefined = không rõ */
  vision?: boolean;
  /** Trần độ dài đầu ra (token), nếu nhà cung cấp có báo */
  maxOutput?: number;
  /** Thời điểm nhà cung cấp phát hành (giây, nếu có) — để chọn bản mới nhất */
  created?: number;
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
          return { id, name: m.displayName ? `${m.displayName} (${id})` : id, context: m.inputTokenLimit, maxOutput: m.outputTokenLimit };
        })
        .filter((m) => !NON_CHAT.test(m.id))
    );
  }
  if (provider.kind === "anthropic" && !provider.dialect) {
    const list: any[] = json?.data || [];
    return sortModels(
      list.map((m) => {
        const t = m.created_at ? Date.parse(m.created_at) : NaN;
        return {
          id: m.id,
          name: m.display_name ? `${m.display_name} (${m.id})` : m.id,
          // Models API (từ 3/2026): max_input_tokens = sức chứa, max_tokens = trần đầu ra
          ...(typeof m.max_input_tokens === "number" ? { context: m.max_input_tokens } : {}),
          ...(typeof m.max_tokens === "number" ? { maxOutput: m.max_tokens } : {}),
          ...(Number.isFinite(t) ? { created: Math.floor(t / 1000) } : {}),
        };
      })
    );
  }
  // Cohere dùng API riêng: { models: [{ name, context_length }] }
  if (Array.isArray(json?.models) && !Array.isArray(json?.data)) {
    return sortModels(
      json.models.map((m: any) => ({ id: m.name || m.id, name: m.name || m.id, context: m.context_length }))
    );
  }
  const only = provider.dialect?.modelFilter ? new RegExp(provider.dialect.modelFilter) : undefined;
  const list: any[] = (json?.data || []).filter((m: any) => !only || only.test(String(m.id)));
  return sortModels(
    list
      .filter((m) => !NON_CHAT.test(String(m.id)))
      .filter((m) => m.active !== false)
      .map((m) => {
        const id = String(m.id);
        const p = m.pricing;
        const free =
          provider.id === "openrouter" || provider.id === "openrouter-claude"
            ? id.endsWith(":free") || (p && Number(p.prompt) === 0 && Number(p.completion) === 0)
            : undefined;
        const sp: any[] | undefined = Array.isArray(m.supported_parameters) ? m.supported_parameters : undefined;
        const mods: any[] | undefined = Array.isArray(m.architecture?.input_modalities) ? m.architecture.input_modalities : undefined;
        return {
          id,
          name: m.name && m.name !== id ? `${m.name} (${id})` : id,
          free: free || undefined,
          context: m.context_length || m.context_window || undefined,
          tools: sp ? sp.includes("tools") : undefined,
          vision: mods ? mods.includes("image") : undefined,
          created: typeof m.created === "number" ? m.created : undefined,
        };
      })
  );
}

export function modelsRequest(provider: Provider, key: string): { url: string; init: RequestInit } {
  if (provider.kind === "gemini") {
    return {
      url: `${provider.baseUrl}/models?pageSize=1000`,
      // Key đi trong tiêu đề, không nằm trong đường dẫn
      init: { method: "GET", headers: { "x-goog-api-key": key } },
    };
  }
  if (provider.kind === "anthropic" && !provider.dialect) {
    return {
      url: `${provider.baseUrl}/models?limit=1000`,
      init: { method: "GET", headers: anthropicHeaders(key) },
    };
  }
  return {
    url: provider.dialect?.modelsUrl || provider.modelsUrl || `${provider.baseUrl}/models`,
    init: { method: "GET", headers: { Authorization: `Bearer ${key}` } },
  };
}

/** Tiêu đề của một yêu cầu kiểu Anthropic. Nguồn có `dialect` gửi key theo cách của nguồn đó; `bare` = không kèm tiêu đề anthropic-*
 *  (bước dò trước của trình duyệt ở nhiều nguồn chỉ cho Authorization và Content-Type — đo 17/11/2026). */
export function anthropicHeaders(key: string, provider?: Provider): Record<string, string> {
  const d = provider?.dialect;
  if (d) {
    const auth: Record<string, string> = d.auth === "bearer" ? { Authorization: `Bearer ${key}` } : { "x-api-key": key };
    return d.bare ? { ...auth, "content-type": "application/json" } : { ...auth, "anthropic-version": "2023-06-01", "content-type": "application/json" };
  }
  return {
    "x-api-key": key,
    "anthropic-version": "2023-06-01",
    "anthropic-dangerous-direct-browser-access": "true",
    "content-type": "application/json",
  };
}

/** Danh sách model cố định của một nguồn kiểu Anthropic (ghi theo tài liệu của nguồn, hoặc do người dùng tự khai) */
export function staticModels(provider: Provider): ModelInfo[] {
  return (provider.dialect?.models || []).map((m) => ({
    id: m.id,
    name: m.name || m.id,
    ...(m.context ? { context: m.context } : {}),
    ...(typeof m.vision === "boolean" ? { vision: m.vision } : {}),
    tools: true,
  }));
}

export async function fetchModels(provider: Provider, key: string, fetchFn: typeof fetch = fetch): Promise<ModelInfo[]> {
  const d = provider.dialect;
  if (d) {
    const fixed = staticModels(provider);
    // Không có địa chỉ liệt kê model: dùng danh sách cố định (key chưa được kiểm ở bước này — sai thì báo ở lượt gọi đầu)
    if (!d.modelsUrl) {
      if (!fixed.length) throw new Error("Nguồn này chưa khai tên model nào.");
      return fixed;
    }
    try {
      const { url, init } = modelsRequest(provider, key);
      const res = await fetchDiagnosed(provider, url, init, fetchFn);
      await ensureOk(res, provider);
      const models = normalizeModels(provider, await res.json());
      if (models.length) return models;
    } catch (e) {
      // Có danh sách cố định thì dùng tạm; không có thì báo lỗi thật
      if (!fixed.length) throw e;
    }
    if (!fixed.length) throw new Error("Nhà cung cấp không trả về model chat nào.");
    return fixed;
  }
  const { url, init } = modelsRequest(provider, key);
  const res = await fetchDiagnosed(provider, url, init, fetchFn);
  await ensureOk(res, provider);
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
