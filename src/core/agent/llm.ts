/** Gọi model một bước, có khai báo công cụ, nhận chữ chạy dần + các lệnh gọi công cụ. 3 kiểu API: OpenAI-compatible, Anthropic, Gemini. */
import { Provider } from "../providers";
import { anthropicHeaders } from "../models";
import { ApiError, ensureOk } from "../errors";
import { createSSEParser } from "../sse";
import { AgentMessage, StepResult, ToolCall, ToolSpec } from "./types";

const ID_CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
/** id 9 ký tự chữ-số: hợp lệ với mọi nhà cung cấp (Mistral bắt buộc đúng 9 ký tự) */
export function newCallId(): string {
  let s = "";
  for (let i = 0; i < 9; i++) s += ID_CHARS[Math.floor(Math.random() * ID_CHARS.length)];
  return s;
}

function parseArgs(raw: any): any {
  if (raw && typeof raw === "object") return raw;
  const s = String(raw ?? "").trim();
  if (!s) return {};
  try {
    return JSON.parse(s);
  } catch {
    // Thử bỏ phần thừa sau dấu } cuối cùng
    const end = s.lastIndexOf("}");
    if (end > 0) {
      try {
        return JSON.parse(s.slice(0, end + 1));
      } catch {
        /* bỏ qua */
      }
    }
    return { __invalid_json: s.slice(0, 500) };
  }
}

/** Gemini chỉ nhận một tập con JSON Schema */
export function toGeminiSchema(schema: any): any {
  if (Array.isArray(schema)) return schema.map(toGeminiSchema);
  if (!schema || typeof schema !== "object") return schema;
  const out: any = {};
  for (const k of ["type", "description", "enum", "required", "nullable", "format"]) {
    if (schema[k] !== undefined) out[k] = schema[k];
  }
  if (Array.isArray(out.type)) out.type = out.type.find((t: string) => t !== "null") || "string";
  if (schema.properties) {
    out.properties = {};
    for (const [k, v] of Object.entries(schema.properties)) out.properties[k] = toGeminiSchema(v);
  }
  if (schema.items) out.items = toGeminiSchema(schema.items);
  if (out.type === "array" && !out.items) out.items = { type: "string" };
  return out;
}

function hasParams(t: ToolSpec): boolean {
  return !!t.parameters && Object.keys(t.parameters.properties || {}).length > 0;
}

/** Tạo yêu cầu HTTP cho một bước — tách riêng để kiểm thử. */
export function buildStepRequest(
  provider: Provider,
  key: string,
  model: string,
  messages: AgentMessage[],
  tools: ToolSpec[],
  opts: { webSearch?: boolean } = {}
): { url: string; init: RequestInit } {
  const system = messages
    .filter((m) => m.role === "system")
    .map((m) => (m as any).content)
    .join("\n\n");
  const turns = messages.filter((m) => m.role !== "system");

  if (provider.kind === "anthropic") {
    const out: any[] = [];
    const pushUser = (blocks: any[]) => {
      const last = out[out.length - 1];
      if (last && last.role === "user") last.content.push(...blocks);
      else out.push({ role: "user", content: blocks });
    };
    for (const m of turns) {
      if (m.role === "user")
        pushUser([
          ...(m.images || []).map((im) => ({ type: "image", source: { type: "base64", media_type: im.mime, data: im.data } })),
          { type: "text", text: m.content || " " },
        ]);
      else if (m.role === "tool")
        pushUser([{ type: "tool_result", tool_use_id: m.toolCallId, content: m.content || "(trống)" }]);
      else if (m.role === "assistant") {
        const blocks: any[] = [];
        if (m.content) blocks.push({ type: "text", text: m.content });
        for (const tc of m.toolCalls || []) blocks.push({ type: "tool_use", id: tc.id, name: tc.name, input: tc.args || {} });
        if (!blocks.length) blocks.push({ type: "text", text: "(trống)" });
        out.push({ role: "assistant", content: blocks });
      }
    }
    const body: any = {
      model,
      max_tokens: /claude-3-(haiku|opus|sonnet)/.test(model) ? 4096 : 8192,
      stream: true,
      messages: out,
    };
    if (system) body.system = system;
    if (tools.length) {
      body.tools = tools.map((t) => ({
        name: t.name,
        description: t.description,
        input_schema: t.parameters || { type: "object", properties: {} },
      }));
    }
    if (opts.webSearch) body.tools = [...(body.tools || []), { type: "web_search_20250305", name: "web_search", max_uses: 5 }];
    return {
      url: `${provider.baseUrl}/messages`,
      init: { method: "POST", headers: anthropicHeaders(key), body: JSON.stringify(body) },
    };
  }

  if (provider.kind === "gemini") {
    // id gốc do Gemini trả về thì phải gửi lại trong functionResponse
    const nativeIds = new Set<string>();
    for (const m of turns) {
      if (m.role === "assistant") for (const tc of m.toolCalls || []) if (tc.nativeId) nativeIds.add(tc.id);
    }
    const contents: any[] = [];
    for (const m of turns) {
      if (m.role === "user")
        contents.push({
          role: "user",
          parts: [...(m.images || []).map((im) => ({ inlineData: { mimeType: im.mime, data: im.data } })), { text: m.content || " " }],
        });
      else if (m.role === "assistant") {
        if (m.raw?.kind === "gemini" && m.raw.parts?.length) {
          contents.push({ role: "model", parts: m.raw.parts });
        } else {
          const parts: any[] = [];
          if (m.content) parts.push({ text: m.content });
          for (const tc of m.toolCalls || []) parts.push({ functionCall: { name: tc.name, args: tc.args || {} } });
          if (!parts.length) parts.push({ text: " " });
          contents.push({ role: "model", parts });
        }
      } else if (m.role === "tool") {
        let response: any;
        try {
          const parsed = JSON.parse(m.content);
          response = parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : { result: parsed };
        } catch {
          response = { result: m.content };
        }
        const part: any = { functionResponse: { name: m.name, response } };
        if (nativeIds.has(m.toolCallId)) part.functionResponse.id = m.toolCallId;
        const prev = contents[contents.length - 1];
        // Gemini: các kết quả của nhiều lệnh gọi song song phải nằm chung một lượt
        if (prev && prev.role === "user" && prev.parts.every((p: any) => p.functionResponse)) prev.parts.push(part);
        else contents.push({ role: "user", parts: [part] });
      }
    }
    const body: any = { contents };
    if (system) body.systemInstruction = { parts: [{ text: system }] };
    if (opts.webSearch && !tools.length) body.tools = [{ google_search: {} }];
    if (tools.length) {
      body.tools = [
        {
          functionDeclarations: tools.map((t) =>
            hasParams(t)
              ? { name: t.name, description: t.description, parameters: toGeminiSchema(t.parameters) }
              : { name: t.name, description: t.description }
          ),
        },
      ];
    }
    return {
      url: `${provider.baseUrl}/models/${encodeURIComponent(model)}:streamGenerateContent?alt=sse&key=${encodeURIComponent(key)}`,
      init: { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) },
    };
  }

  // OpenAI-compatible
  const msgs: any[] = [];
  if (system) msgs.push({ role: "system", content: system });
  for (const m of turns) {
    if (m.role === "user") {
      if (m.images?.length) {
        msgs.push({
          role: "user",
          content: [
            { type: "text", text: m.content },
            ...m.images.map((im) => ({ type: "image_url", image_url: { url: `data:${im.mime};base64,${im.data}` } })),
          ],
        });
      } else msgs.push({ role: "user", content: m.content });
    } else if (m.role === "tool") msgs.push({ role: "tool", tool_call_id: m.toolCallId, content: m.content });
    else if (m.role === "assistant") {
      const calls = m.toolCalls || [];
      const msg: any = { role: "assistant", content: m.content || (calls.length ? null : "") };
      if (calls.length) {
        msg.tool_calls = calls.map((tc) => ({
          id: tc.id,
          type: "function",
          function: { name: tc.name, arguments: JSON.stringify(tc.args || {}) },
        }));
      }
      msgs.push(msg);
    }
  }
  const headers: Record<string, string> = { Authorization: `Bearer ${key}`, "content-type": "application/json" };
  if (provider.id === "openrouter") headers["X-Title"] = "AI Da Nguon Office Add-in";
  const body: any = { model, messages: msgs, stream: true };
  if (opts.webSearch && provider.id === "openrouter") body.plugins = [{ id: "web", max_results: 5 }];
  if (tools.length) {
    body.tools = tools.map((t) => ({
      type: "function",
      function: { name: t.name, description: t.description, parameters: t.parameters || { type: "object", properties: {} } },
    }));
  }
  return { url: `${provider.baseUrl}/chat/completions`, init: { method: "POST", headers, body: JSON.stringify(body) } };
}

/** Gom các sự kiện luồng thành chữ + lệnh gọi công cụ. */
export function createStepAccumulator(kind: string, onText: (d: string) => void) {
  let text = "";
  let finishReason: string | undefined;
  // OpenAI
  const oaCalls: { id: string; name: string; args: string }[] = [];
  let lastKey = -1;
  // Anthropic
  const blocks: Record<number, { type: string; id?: string; name?: string; json: string }> = {};
  // Gemini
  const gParts: any[] = [];
  const gCalls: ToolCall[] = [];

  const addText = (d: string) => {
    if (!d) return;
    text += d;
    onText(d);
  };

  return {
    push(data: string) {
      let j: any;
      try {
        j = JSON.parse(data);
      } catch {
        return;
      }
      if (j?.error) {
        const msg = j.error.message || (typeof j.error === "string" ? j.error : JSON.stringify(j.error));
        throw new ApiError(`Nhà cung cấp báo lỗi: ${msg}`);
      }
      if (kind === "anthropic") {
        if (j.type === "content_block_start") {
          const b = j.content_block || {};
          blocks[j.index] = { type: b.type, id: b.id, name: b.name, json: "" };
          if (b.type === "text" && b.text) addText(b.text);
        } else if (j.type === "content_block_delta") {
          const d = j.delta || {};
          if (d.type === "text_delta") addText(d.text || "");
          else if (d.type === "input_json_delta") {
            const b = blocks[j.index] || (blocks[j.index] = { type: "tool_use", json: "" });
            b.json += d.partial_json || "";
          }
        } else if (j.type === "message_delta") {
          finishReason = j.delta?.stop_reason || finishReason;
        }
        return;
      }
      if (kind === "gemini") {
        const cand = j?.candidates?.[0];
        if (!cand) return;
        for (const p of cand.content?.parts || []) {
          gParts.push(p);
          if (p.functionCall) {
            gCalls.push({
              id: p.functionCall.id || newCallId(),
              name: p.functionCall.name,
              args: p.functionCall.args || {},
              nativeId: !!p.functionCall.id,
            });
          } else if (typeof p.text === "string" && !p.thought) addText(p.text);
        }
        if (cand.finishReason) finishReason = cand.finishReason;
        return;
      }
      const choice = j?.choices?.[0];
      if (!choice) return;
      const delta = choice.delta || choice.message || {};
      if (typeof delta.content === "string") addText(delta.content);
      for (const tc of delta.tool_calls || []) {
        let key: number;
        if (typeof tc.index === "number") key = tc.index;
        else if (tc.id && !oaCalls.some((c) => c.id === tc.id)) key = oaCalls.length;
        else key = lastKey >= 0 ? lastKey : 0;
        lastKey = key;
        const entry = oaCalls[key] || (oaCalls[key] = { id: "", name: "", args: "" });
        if (tc.id) entry.id = tc.id;
        if (tc.function?.name && !entry.name) entry.name = tc.function.name;
        const a = tc.function?.arguments;
        if (a !== undefined && a !== null) entry.args += typeof a === "string" ? a : JSON.stringify(a);
      }
      if (choice.finish_reason) finishReason = choice.finish_reason;
    },
    result(): StepResult {
      if (kind === "anthropic") {
        const calls: ToolCall[] = Object.keys(blocks)
          .map(Number)
          .sort((a, b) => a - b)
          .map((i) => blocks[i])
          .filter((b) => b.type === "tool_use" && b.name)
          .map((b) => ({ id: b.id || newCallId(), name: b.name!, args: parseArgs(b.json) }));
        return { text, toolCalls: calls, finishReason };
      }
      if (kind === "gemini") {
        return { text, toolCalls: gCalls, raw: { kind: "gemini", parts: gParts }, finishReason };
      }
      const calls = oaCalls
        .filter((c) => c && c.name)
        .map((c) => ({ id: c.id || newCallId(), name: c.name, args: parseArgs(c.args) }));
      return { text, toolCalls: calls, finishReason };
    },
  };
}

export function isToolUnsupportedError(e: any): boolean {
  const msg = String(e?.message || "");
  const status = e?.status;
  return (
    (status === 400 || status === 404 || status === 422 || status === undefined) &&
    /(tool|function[ _-]?call|tool_choice|does not support tools|No endpoints found that support)/i.test(msg)
  );
}

export async function callModelStep(opts: {
  provider: Provider;
  key: string;
  model: string;
  messages: AgentMessage[];
  tools: ToolSpec[];
  signal?: AbortSignal;
  onText: (d: string) => void;
  fetchFn?: typeof fetch;
  webSearch?: boolean;
}): Promise<StepResult> {
  const { url, init } = buildStepRequest(opts.provider, opts.key, opts.model, opts.messages, opts.tools, {
    webSearch: opts.webSearch,
  });
  const fetchFn = opts.fetchFn || fetch;
  const res = await fetchFn(url, { ...init, signal: opts.signal });
  await ensureOk(res);
  const acc = createStepAccumulator(opts.provider.kind, opts.onText);
  const parser = createSSEParser((d) => acc.push(d));
  if (res.body) {
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      parser.push(decoder.decode(value, { stream: true }));
    }
  } else {
    parser.push(await res.text());
  }
  parser.end();
  return acc.result();
}
