/** Gửi hội thoại và nhận câu trả lời chạy dần (streaming) cho 3 kiểu API. */
import { Provider } from "./providers";
import { anthropicHeaders } from "./models";
import { ensureOk } from "./errors";
import { readSSE } from "./sse";

export interface ChatMessage {
  role: "user" | "assistant" | "system";
  content: string;
}

export interface ChatRequest {
  provider: Provider;
  key: string;
  model: string;
  messages: ChatMessage[];
  signal?: AbortSignal;
  onDelta: (text: string) => void;
  fetchFn?: typeof fetch;
}

/** Tạo yêu cầu HTTP — tách riêng để kiểm thử. */
export function buildChatRequest(r: Omit<ChatRequest, "onDelta" | "signal" | "fetchFn">): { url: string; init: RequestInit } {
  const { provider, key, model, messages } = r;
  const system = messages.filter((m) => m.role === "system").map((m) => m.content).join("\n\n");
  const turns = messages.filter((m) => m.role !== "system");

  if (provider.kind === "gemini") {
    const body: any = {
      contents: turns.map((m) => ({ role: m.role === "assistant" ? "model" : "user", parts: [{ text: m.content }] })),
    };
    if (system) body.systemInstruction = { parts: [{ text: system }] };
    return {
      url: `${provider.baseUrl}/models/${encodeURIComponent(model)}:streamGenerateContent?alt=sse&key=${encodeURIComponent(key)}`,
      init: { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) },
    };
  }
  if (provider.kind === "anthropic") {
    const body: any = { model, max_tokens: 8192, stream: true, messages: turns };
    if (system) body.system = system;
    return {
      url: `${provider.baseUrl}/messages`,
      init: { method: "POST", headers: anthropicHeaders(key), body: JSON.stringify(body) },
    };
  }
  const headers: Record<string, string> = { Authorization: `Bearer ${key}`, "content-type": "application/json" };
  if (provider.id === "openrouter") {
    headers["X-Title"] = "AI Da Nguon Office Add-in";
  }
  return {
    url: `${provider.baseUrl}/chat/completions`,
    init: { method: "POST", headers, body: JSON.stringify({ model, messages, stream: true }) },
  };
}

/** Lấy phần chữ mới từ một sự kiện SSE. */
export function extractDelta(provider: Provider, data: string): string {
  let j: any;
  try {
    j = JSON.parse(data);
  } catch {
    return "";
  }
  if (j?.error) throw new Error(j.error.message || JSON.stringify(j.error));
  if (provider.kind === "gemini") {
    const parts: any[] = j?.candidates?.[0]?.content?.parts || [];
    return parts.map((p) => (p.thought ? "" : p.text || "")).join("");
  }
  if (provider.kind === "anthropic") {
    if (j?.type === "content_block_delta" && j.delta?.type === "text_delta") return j.delta.text || "";
    return "";
  }
  return j?.choices?.[0]?.delta?.content || "";
}

export async function streamChat(r: ChatRequest): Promise<string> {
  const { url, init } = buildChatRequest(r);
  const fetchFn = r.fetchFn || fetch;
  const res = await fetchFn(url, { ...init, signal: r.signal });
  await ensureOk(res);
  let full = "";
  if (!res.body) {
    // Một số môi trường không hỗ trợ luồng: đọc một lần
    const text = await res.text();
    text.split("\n").forEach((line) => {
      if (line.startsWith("data:")) {
        const d = extractDelta(r.provider, line.slice(5).trim());
        full += d;
        if (d) r.onDelta(d);
      }
    });
    return full;
  }
  await readSSE(res.body, (data) => {
    const d = extractDelta(r.provider, data);
    if (d) {
      full += d;
      r.onDelta(d);
    }
  });
  return full;
}
