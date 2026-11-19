/** Gọi model một bước, có khai báo công cụ, nhận chữ chạy dần + các lệnh gọi công cụ. 3 kiểu API: OpenAI-compatible, Anthropic, Gemini. */
import { Provider } from "../providers";
import { anthropicHeaders } from "../models";
import { ApiError, ensureOk, explainApiError, fetchDiagnosed, isRetryable, maskSecrets, rememberSecret } from "../errors";
import { createSSEParser } from "../sse";
import { AgentMessage, RawAssistant, StepResult, ToolCall, ToolSpec, Usage } from "./types";
import { prepareHistory } from "./history";
import { dialectThinkingParams, geminiThinkingConfig, isThinkingParamError } from "./thinking";

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
  // Gemini chỉ nhận enum là chuỗi trên trường kiểu string: enum số (ví dụ cỡ chữ [13, 14]) làm cả yêu cầu bị từ chối 400
  // (đo 16/11/2026 với mọi model Gemini trong Word) → bỏ enum, ghi các giá trị cho phép vào mô tả.
  if (Array.isArray(out.enum) && (out.type !== "string" || out.enum.some((v: unknown) => typeof v !== "string"))) {
    out.description = `${out.description ? out.description + " " : ""}Allowed values: ${out.enum.join(", ")}.`;
    delete out.enum;
  }
  if (schema.properties) {
    out.properties = {};
    for (const [k, v] of Object.entries(schema.properties)) out.properties[k] = toGeminiSchema(v);
  }
  if (schema.items) out.items = toGeminiSchema(schema.items);
  if (out.type === "array" && !out.items) out.items = { type: "string" };
  return out;
}

/** Trần độ dài đầu ra gửi cho Anthropic. Biết trần của model (Models API, trường max_tokens) thì dùng, tối đa 32.000;
 *  không biết thì theo đời model: Claude 3 là 4.096, Claude 3.5 là 8.192, còn lại 16.000 (mọi model từ 3.7 đều nhận). */
export function anthropicMaxTokens(model: string, modelMax?: number): number {
  if (modelMax && modelMax > 0) return Math.min(Math.floor(modelMax), 32000);
  if (/claude-3-(haiku|opus|sonnet)/.test(model)) return 4096;
  if (/claude-3-5/.test(model)) return 8192;
  return 16000;
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
  opts: {
    webSearch?: boolean;
    vision?: boolean;
    maxOutput?: number;
    /** Công tắc "Suy nghĩ kỹ" */
    think?: boolean;
    /** Không gửi tham số suy nghĩ và không gửi lại khối suy nghĩ (lần gửi lại sau khi nhà cung cấp từ chối các phần đó) */
    plain?: boolean;
    /** Gemini: lệnh gọi công cụ do nguồn khác sinh ra (không có phần gốc của Gemini, nên không có chữ ký suy nghĩ) gửi dưới dạng chữ,
     *  kèm kết quả của nó — xoay vòng nguồn miễn phí, Đợt 9 R3. Chỉ đổi yêu cầu gửi đi, không đổi lịch sử. */
    foreignToolsAsText?: boolean;
  } = {}
): { url: string; init: RequestInit } {
  const system = messages
    .filter((m) => m.role === "system")
    .map((m) => (m as any).content)
    .join("\n\n");
  // Chuẩn hóa cho nguồn đích: mã lời gọi công cụ, cặp gọi–kết quả, ảnh cũ, chữ trắng (đổi nguồn giữa hội thoại)
  const turns = prepareHistory(messages, { provider, vision: opts.vision });
  const nonBlank = (t: string) => !!t && !!t.trim();

  if (provider.kind === "anthropic") {
    const out: any[] = [];
    const pushUser = (blocks: any[]) => {
      const last = out[out.length - 1];
      if (last && last.role === "user") last.content.push(...blocks);
      else out.push({ role: "user", content: blocks });
    };
    for (const m of turns) {
      if (m.role === "user") {
        // Tệp PDF đứng trước, rồi ảnh, rồi chữ (tài liệu chính thức: khối document đặt trước khối chữ)
        const imgs = [
          ...(m.files || []).map((f) => ({ type: "document", source: { type: "base64", media_type: f.mime, data: f.data } })),
          ...(m.images || []).map((im) => ({ type: "image", source: { type: "base64", media_type: im.mime, data: im.data } })),
        ];
        // Anthropic từ chối khối chữ rỗng hoặc chỉ có khoảng trắng
        pushUser(nonBlank(m.content) ? [...imgs, { type: "text", text: m.content }] : imgs.length ? imgs : [{ type: "text", text: "(trống)" }]);
      } else if (m.role === "tool")
        pushUser([{ type: "tool_result", tool_use_id: m.toolCallId, content: nonBlank(m.content) ? m.content : "(trống)" }]);
      else if (m.role === "assistant") {
        const blocks: any[] = [];
        const calls = (m.toolCalls || []).slice();
        const useCall = () => {
          const tc = calls.shift();
          if (tc) blocks.push({ type: "tool_use", id: tc.id, name: tc.name, input: tc.args || {} });
        };
        let textDone = false;
        const useText = () => {
          if (!textDone && nonBlank(m.content)) blocks.push({ type: "text", text: m.content });
          textDone = true;
        };
        // Câu trả lời có khối suy nghĩ: gửi lại đúng thứ tự gốc, khối suy nghĩ nguyên văn kèm chữ ký (tài liệu: "pass thinking
        // blocks back unmodified"). Cách dựng này cho cùng một kết quả ở mọi lần gửi nên phần đầu hội thoại không đổi.
        if (!opts.plain && m.raw?.kind === "anthropic" && Array.isArray(m.raw.blocks)) {
          for (const b of m.raw.blocks) {
            if (b?.type === "thinking" || b?.type === "redacted_thinking") blocks.push(b);
            else if (b?.type === "text") {
              // Mỗi khối chữ gửi lại đúng vị trí và đúng nguyên văn của nó (câu trả lời có thể xen chữ giữa các khối suy nghĩ:
              // gộp chữ lại một chỗ là đổi phần đứng trước khối suy nghĩ sau — soát mã độc lập 15/11/2026, mục M5)
              if (typeof b.text === "string") {
                if (nonBlank(b.text)) blocks.push({ type: "text", text: b.text });
                textDone = true;
              } else useText();
            } else if (b?.type === "tool_use") {
              if (!m.raw.blocks.some((x: any) => x?.type === "text" && typeof x.text === "string")) useText();
              useCall();
            }
          }
        }
        useText();
        while (calls.length) useCall();
        if (!blocks.length) blocks.push({ type: "text", text: "(trống)" });
        out.push({ role: "assistant", content: blocks });
      }
    }
    // Bộ nhớ đệm lời nhắc: một mốc ở cuối lời nhắc hệ thống (phủ cả danh sách công cụ, đứng trước trong thứ tự dựng),
    // một mốc ở khối cuối của tin nhắn cuối (phủ phần hội thoại đã có). Lịch sử chỉ thêm nên lượt sau đọc lại được phần đầu.
    const CACHE = { type: "ephemeral" };
    const d = provider.dialect;
    // Nguồn bỏ qua cache_control (ghi trong phương ngữ) thì không gắn mốc
    const useCache = d?.cache !== false;
    const lastMsg = out[out.length - 1];
    if (useCache && lastMsg?.content?.length) {
      const i = lastMsg.content.length - 1;
      lastMsg.content[i] = { ...lastMsg.content[i], cache_control: CACHE };
    }
    const maxTokens = anthropicMaxTokens(model, opts.maxOutput);
    const body: any = {
      model,
      max_tokens: maxTokens,
      stream: true,
      messages: out,
    };
    // Tham số suy nghĩ nằm ngoài system / tools / messages: bật hay tắt không làm vỡ bộ nhớ đệm lời nhắc
    const think = opts.plain ? {} : dialectThinkingParams(provider, model, !!opts.think, maxTokens);
    if (think.thinking) body.thinking = think.thinking;
    if (think.output_config) body.output_config = think.output_config;
    if (system) body.system = [useCache ? { type: "text", text: system, cache_control: CACHE } : { type: "text", text: system }];
    if (tools.length) {
      body.tools = tools.map((t) => ({
        name: t.name,
        description: t.description,
        input_schema: t.parameters || { type: "object", properties: {} },
      }));
    }
    if (opts.webSearch && (!d || d.webSearch)) body.tools = [...(body.tools || []), { type: "web_search_20250305", name: "web_search", max_uses: 5 }];
    return {
      url: `${provider.baseUrl}/messages`,
      init: {
        method: "POST",
        headers: think.beta ? { ...anthropicHeaders(key, provider), "anthropic-beta": think.beta } : anthropicHeaders(key, provider),
        body: JSON.stringify(body),
      },
    };
  }

  if (provider.kind === "gemini") {
    // id gốc do Gemini trả về thì phải gửi lại trong functionResponse
    const nativeIds = new Set<string>();
    for (const m of turns) {
      if (m.role === "assistant") for (const tc of m.toolCalls || []) if (tc.nativeId) nativeIds.add(tc.id);
    }
    /** Lệnh gọi của nguồn khác gửi thành chữ (Đợt 9 R3) */
    const foreign = new Set<string>();
    if (opts.foreignToolsAsText)
      for (const m of turns)
        if (m.role === "assistant" && !(m.raw?.kind === "gemini" && m.raw.parts?.length)) for (const tc of m.toolCalls || []) foreign.add(tc.id);
    const contents: any[] = [];
    for (const m of turns) {
      if (m.role === "user")
        contents.push({
          role: "user",
          parts: [
            ...(m.files || []).map((f) => ({ inlineData: { mimeType: f.mime, data: f.data } })),
            ...(m.images || []).map((im) => ({ inlineData: { mimeType: im.mime, data: im.data } })),
            { text: m.content || " " },
          ],
        });
      else if (m.role === "assistant") {
        if (m.raw?.kind === "gemini" && m.raw.parts?.length) {
          contents.push({ role: "model", parts: m.raw.parts });
        } else {
          const parts: any[] = [];
          if (m.content) parts.push({ text: m.content });
          for (const tc of m.toolCalls || [])
            parts.push(
              foreign.has(tc.id)
                ? { text: `[Bước do model trước thực hiện: đã gọi công cụ ${tc.name} với tham số ${JSON.stringify(tc.args || {})}]` }
                : { functionCall: { name: tc.name, args: tc.args || {} } }
            );
          if (!parts.length) parts.push({ text: " " });
          contents.push({ role: "model", parts });
        }
      } else if (m.role === "tool" && foreign.has(m.toolCallId)) {
        const part = { text: `[Kết quả công cụ ${m.name} (bước do model trước thực hiện)]: ${m.content || "(trống)"}` };
        const prev = contents[contents.length - 1];
        if (prev && prev.role === "user" && !prev.parts.some((p: any) => p.functionResponse)) prev.parts.push(part);
        else contents.push({ role: "user", parts: [part] });
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
    const thinkingConfig = opts.think && !opts.plain ? geminiThinkingConfig(model) : undefined;
    if (thinkingConfig) body.generationConfig = { thinkingConfig };
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
    // Key đi trong tiêu đề, không nằm trong đường dẫn (đường dẫn dễ lọt vào nhật ký, thông báo lỗi)
    return {
      url: `${provider.baseUrl}/models/${encodeURIComponent(model)}:streamGenerateContent?alt=sse`,
      init: { method: "POST", headers: { "content-type": "application/json", "x-goog-api-key": key }, body: JSON.stringify(body) },
    };
  }

  // OpenAI-compatible
  let hasFiles = false;
  const msgs: any[] = [];
  if (system) msgs.push({ role: "system", content: system });
  for (const m of turns) {
    if (m.role === "user") {
      if (m.images?.length || m.files?.length) {
        if (m.files?.length) hasFiles = true;
        msgs.push({
          role: "user",
          content: [
            { type: "text", text: nonBlank(m.content) ? m.content : "(trống)" },
            ...(m.images || []).map((im) => ({ type: "image_url", image_url: { url: `data:${im.mime};base64,${im.data}` } })),
            ...(m.files || []).map((f) => ({ type: "file", file: { filename: f.name, file_data: `data:${f.mime};base64,${f.data}` } })),
          ],
        });
      } else msgs.push({ role: "user", content: nonBlank(m.content) ? m.content : "(trống)" });
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
      // OpenRouter: gửi lại nguyên các khối suy nghĩ của câu trả lời trước (tài liệu: bắt buộc khi tiếp tục vòng lặp công cụ)
      if (!opts.plain && provider.id === "openrouter" && m.raw?.kind === "openrouter" && m.raw.reasoning?.length) msg.reasoning_details = m.raw.reasoning;
      msgs.push(msg);
    }
  }
  const headers: Record<string, string> = { Authorization: `Bearer ${key}`, "content-type": "application/json" };
  if (provider.id === "openrouter") headers["X-Title"] = "AI Da Nguon Office Add-in";
  const body: any = { model, messages: msgs, stream: true };
  if (opts.webSearch && provider.id === "openrouter") body.plugins = [{ id: "web", max_results: 5 }];
  // OpenRouter: chỉ rõ để model tự đọc tệp PDF; không để OpenRouter tự chuyển sang bộ OCR tính phí riêng
  if (hasFiles && provider.id === "openrouter") body.plugins = [...(body.plugins || []), { id: "file-parser", pdf: { engine: "native" } }];
  if (opts.think && !opts.plain) {
    if (provider.id === "openrouter") body.reasoning = { effort: "high" };
    else if (provider.id === "openai") body.reasoning_effort = "high";
  }
  if (tools.length) {
    body.tools = tools.map((t) => ({
      type: "function",
      function: { name: t.name, description: t.description, parameters: t.parameters || { type: "object", properties: {} } },
    }));
  }
  return { url: `${provider.baseUrl}/chat/completions`, init: { method: "POST", headers, body: JSON.stringify(body) } };
}

/** Gom các sự kiện luồng thành chữ + lệnh gọi công cụ. */
export function createStepAccumulator(kind: string, onText: (d: string) => void, onThinking?: () => void, provider?: Provider) {
  let thinkingSeen = false;
  const noteThinking = () => {
    if (thinkingSeen) return;
    thinkingSeen = true;
    onThinking?.();
  };
  // OpenRouter: các khối suy nghĩ (reasoning_details) tới theo từng mẩu, ghép theo index
  const orDetails: any[] = [];
  const mergeDetail = (d: any) => {
    if (!d || typeof d !== "object") return;
    const at = typeof d.index === "number" ? orDetails.findIndex((x) => x.index === d.index && x.type === d.type) : -1;
    if (at < 0) {
      orDetails.push({ ...d });
      return;
    }
    const cur = orDetails[at];
    for (const [k, v] of Object.entries(d)) {
      if (k === "text" || k === "summary" || k === "data") cur[k] = (typeof cur[k] === "string" ? cur[k] : "") + (typeof v === "string" ? v : "");
      else if (v !== undefined && v !== null && v !== "") cur[k] = v;
    }
  };
  let text = "";
  let finishReason: string | undefined;
  let usage: Usage | undefined;
  const num = (v: any): number | undefined => (typeof v === "number" && Number.isFinite(v) ? v : undefined);
  // Anthropic báo token rải ở message_start (đầu vào) và message_delta (đầu ra): gom dần
  const anth: { in?: number; out?: number; cr?: number; cw?: number } = {};
  const mergeAnthropic = (u: any) => {
    if (!u) return;
    anth.in = num(u.input_tokens) ?? anth.in;
    anth.out = num(u.output_tokens) ?? anth.out;
    anth.cr = num(u.cache_read_input_tokens) ?? anth.cr;
    anth.cw = num(u.cache_creation_input_tokens) ?? anth.cw;
    // input_tokens của Anthropic KHÔNG gồm phần đệm → cộng lại để "input" là tổng đầu vào
    usage = { input: (anth.in || 0) + (anth.cr || 0) + (anth.cw || 0), output: anth.out, cacheRead: anth.cr, cacheWrite: anth.cw };
  };
  // OpenAI
  const oaCalls: { id: string; name: string; args: string }[] = [];
  let lastKey = -1;
  // Anthropic
  const blocks: Record<number, { type: string; id?: string; name?: string; json: string; thinking?: string; signature?: string; data?: string; text?: string }> = {};
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
        // Lỗi giữa luồng không có mã HTTP: suy ra từ loại lỗi để biết có nên tự thử lại
        const type = String(j.error.type || "");
        const code = Number(j.error.code ?? j.error.status);
        const status =
          type === "overloaded_error" ? 529 : type === "rate_limit_error" ? 429 : type === "api_error" ? 500 : code >= 400 && code < 600 ? code : undefined;
        // Lỗi về key hoặc tiền báo giữa luồng (OpenRouter: hết credit sau khi đã trả 200): dùng câu của lý do, không thử lại
        const x = explainApiError(status, String(msg), { provider, body: j });
        if (x.reason) throw new ApiError(x.text, status, undefined, x.reason);
        const err = new ApiError(`Nhà cung cấp báo lỗi: ${maskSecrets(String(msg))}`, status);
        if (x.transient) err.transient = true;
        throw err;
      }
      if (kind === "anthropic") {
        if (j.type === "message_start") {
          mergeAnthropic(j.message?.usage);
        } else if (j.type === "content_block_start") {
          const b = j.content_block || {};
          blocks[j.index] = { type: b.type, id: b.id, name: b.name, json: "" };
          if (b.type === "text") blocks[j.index].text = b.text || "";
          if (b.type === "text" && b.text) addText(b.text);
          if (b.type === "thinking" || b.type === "redacted_thinking") {
            noteThinking();
            blocks[j.index].thinking = b.thinking || "";
            blocks[j.index].signature = b.signature || "";
            blocks[j.index].data = b.data;
          }
        } else if (j.type === "content_block_delta") {
          const d = j.delta || {};
          if (d.type === "text_delta") {
            addText(d.text || "");
            const b = blocks[j.index];
            if (b && b.type === "text") b.text = (b.text || "") + (d.text || "");
          }
          else if (d.type === "input_json_delta") {
            const b = blocks[j.index] || (blocks[j.index] = { type: "tool_use", json: "" });
            b.json += d.partial_json || "";
          } else if (d.type === "thinking_delta") {
            const b = blocks[j.index];
            if (b) b.thinking = (b.thinking || "") + (d.thinking || "");
          } else if (d.type === "signature_delta") {
            const b = blocks[j.index];
            if (b) b.signature = (b.signature || "") + (d.signature || "");
          }
        } else if (j.type === "message_delta") {
          finishReason = j.delta?.stop_reason || finishReason;
          mergeAnthropic(j.usage);
        }
        return;
      }
      if (kind === "gemini") {
        const um = j?.usageMetadata;
        if (um) {
          usage = {
            input: num(um.promptTokenCount),
            output: (num(um.candidatesTokenCount) || 0) + (num(um.thoughtsTokenCount) || 0) || undefined,
            cacheRead: num(um.cachedContentTokenCount),
          };
        }
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
          else if (p.thought) noteThinking();
        }
        if (cand.finishReason) finishReason = cand.finishReason;
        return;
      }
      // Khối cuối của nhiều nguồn kiểu OpenAI chỉ có usage, không có choices (Groq đặt trong x_groq)
      const u = j?.usage || j?.x_groq?.usage;
      if (u && (num(u.prompt_tokens) !== undefined || num(u.completion_tokens) !== undefined)) {
        usage = {
          input: num(u.prompt_tokens),
          output: num(u.completion_tokens),
          cacheRead: num(u.prompt_tokens_details?.cached_tokens) ?? num(u.prompt_cache_hit_tokens),
        };
      }
      const choice = j?.choices?.[0];
      if (!choice) return;
      const delta = choice.delta || choice.message || {};
      if (Array.isArray(delta.reasoning_details) && delta.reasoning_details.length) {
        noteThinking();
        delta.reasoning_details.forEach(mergeDetail);
      } else if (typeof delta.reasoning === "string" && delta.reasoning) noteThinking();
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
        const ordered = Object.keys(blocks)
          .map(Number)
          .sort((a, b) => a - b)
          .map((i) => blocks[i]);
        const calls: ToolCall[] = ordered.filter((b) => b.type === "tool_use" && b.name).map((b) => ({ id: b.id || newCallId(), name: b.name!, args: parseArgs(b.json) }));
        // Giữ khối suy nghĩ (nguyên văn, kèm chữ ký) cùng thứ tự các khối để gửi lại ở bước sau. Khối suy nghĩ chưa có chữ ký
        // (câu trả lời bị cắt giữa chừng) thì bỏ: gửi lại sẽ bị từ chối.
        const rawBlocks: any[] = [];
        for (const b of ordered) {
          if (b.type === "thinking") {
            if (b.signature) rawBlocks.push({ type: "thinking", thinking: b.thinking || "", signature: b.signature });
          } else if (b.type === "redacted_thinking") {
            if (b.data) rawBlocks.push({ type: "redacted_thinking", data: b.data });
          } else if (b.type === "text") rawBlocks.push({ type: "text", text: b.text || "" });
          else if (b.type === "tool_use" && b.name) rawBlocks.push({ type: "tool_use" });
        }
        const raw: RawAssistant | undefined = rawBlocks.some((b) => b.type === "thinking" || b.type === "redacted_thinking") ? { kind: "anthropic", blocks: rawBlocks } : undefined;
        // Luồng đứt trước khi có lý do kết thúc mà lệnh công cụ cuối chưa có tham số: coi như bị cắt dở (không chạy lệnh với tham số rỗng)
        const lastTool = [...ordered].reverse().find((b) => b.type === "tool_use" && b.name);
        const cutOff = !finishReason && !!lastTool && !String(lastTool.json || "").trim();
        return { text, toolCalls: calls, ...(raw ? { raw } : {}), finishReason, truncated: finishReason === "max_tokens" || finishReason === "model_context_window_exceeded" || cutOff, usage };
      }
      if (kind === "gemini") {
        return { text, toolCalls: gCalls, raw: { kind: "gemini", parts: gParts }, finishReason, truncated: finishReason === "MAX_TOKENS", usage };
      }
      const calls = oaCalls
        .filter((c) => c && c.name)
        .map((c) => ({ id: c.id || newCallId(), name: c.name, args: parseArgs(c.args) }));
      const raw: RawAssistant | undefined = orDetails.length ? { kind: "openrouter", reasoning: orDetails } : undefined;
      const lastOa = [...oaCalls].reverse().find((c) => c && c.name);
      const cutOffOa = !finishReason && !!lastOa && !String(lastOa.args || "").trim();
      return { text, toolCalls: calls, ...(raw ? { raw } : {}), finishReason, truncated: finishReason === "length" || cutOffOa, usage };
    },
  };
}

/** Lỗi do model không hỗ trợ gọi công cụ. Chỉ khớp câu báo "không hỗ trợ", không khớp lỗi về mã lời gọi hay tham số công cụ
 *  (trước đây mọi lỗi có chữ "tool" đều bị báo nhầm là "model không gọi được công cụ"). */
const NO_TOOL_SUPPORT =
  /(does ?n[o']t support (tools|tool[ _-]?(use|calling)|function[ _-]?calling)|tools? (is|are) not supported|(tool[ _-]?(use|calling)|function[ _-]?calling)[`'"]? (is )?not (supported|enabled|available)|not support(ed)? (for )?(tool[ _-]?(use|calling)|function[ _-]?calling)|No endpoints found that support tool|tool_choice[^.]*(not supported|unsupported))/i;

export function isToolUnsupportedError(e: any): boolean {
  const msg = String(e?.message || "");
  const status = e?.status;
  return (status === 400 || status === 404 || status === 422 || status === undefined) && NO_TOOL_SUPPORT.test(msg);
}

/** Khoảng chờ trước mỗi lần tự thử lại (mili giây). Số phần tử = số lần thử lại tối đa. */
export const RETRY_DELAYS_MS = [2000, 6000, 15000];
/** Nhà cung cấp bảo chờ lâu hơn mức này (giây) thì không tự chờ, báo người dùng luôn. */
export const MAX_RETRY_AFTER_S = 60;

/** Không có byte nào từ nhà cung cấp trong ngần này mili giây thì coi là treo và hủy yêu cầu. */
export const STALL_MS = 150000;

/** Chờ có thể hủy bằng nút Dừng. */
export function sleepAbortable(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) return reject(new DOMException("Đã dừng", "AbortError"));
    const onAbort = () => {
      clearTimeout(t);
      reject(new DOMException("Đã dừng", "AbortError"));
    };
    const t = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

export interface RetryInfo {
  /** Lần thử lại thứ mấy (từ 1) */
  attempt: number;
  max: number;
  waitMs: number;
  reason: string;
}

export interface CallModelOptions {
  provider: Provider;
  key: string;
  model: string;
  messages: AgentMessage[];
  tools: ToolSpec[];
  signal?: AbortSignal;
  onText: (d: string) => void;
  fetchFn?: typeof fetch;
  webSearch?: boolean;
  /** Model đích nhận ảnh (true/false/không rõ) */
  vision?: boolean;
  /** Trần độ dài đầu ra của model nếu biết (Anthropic) */
  maxOutput?: number;
  /** Công tắc "Suy nghĩ kỹ" */
  think?: boolean;
  /** Trạng thái dùng chung của một lượt: nhà cung cấp đã từ chối tham số suy nghĩ thì các bước sau gửi thẳng không kèm */
  state?: { plain?: boolean };
  /** Model bắt đầu suy nghĩ (chưa có chữ trả lời) */
  onThinking?: () => void;
  /** Nhà cung cấp từ chối tham số suy nghĩ hoặc khối suy nghĩ gửi lại: đã gửi lại không kèm các phần đó */
  onThinkingOff?: (reason: string) => void;
  /** Được gọi trước mỗi lần chờ để thử lại */
  onRetry?: (info: RetryInfo) => void;
  /** Thay bảng khoảng chờ (kiểm thử); mảng rỗng = không thử lại */
  retryDelaysMs?: number[];
  sleep?: (ms: number, signal?: AbortSignal) => Promise<void>;
  /** Thay hạn chờ im lặng STALL_MS (chế độ Miễn phí: 60 giây; kiểm thử) */
  stallMs?: number;
  /** Gemini: lệnh gọi công cụ của nguồn khác gửi thành chữ (xoay vòng, Đợt 9) */
  foreignToolsAsText?: boolean;
}

/** Gọi model một bước; tự thử lại khi nhà cung cấp quá tải hoặc hết hạn mức tạm thời (429, 5xx, 529),
 *  miễn là chưa có chữ nào hiện ra cho người dùng (tránh lặp chữ) và người dùng chưa bấm Dừng. */
export async function callModelStep(input: CallModelOptions): Promise<StepResult> {
  rememberSecret(input.key);
  const opts: CallModelOptions = input.state ? input : { ...input, state: {} };
  const delays = opts.retryDelaysMs ?? RETRY_DELAYS_MS;
  for (let attempt = 0; ; attempt++) {
    let emitted = false;
    try {
      return await callModelOnce({
        ...opts,
        onText: (d) => {
          emitted = true;
          opts.onText(d);
        },
      });
    } catch (e: any) {
      // Nhà cung cấp không nhận tham số suy nghĩ, hoặc từ chối khối suy nghĩ gửi lại (tài liệu Anthropic: bỏ hết khối suy nghĩ
      // rồi gửi lại một lần): gửi lại không kèm các phần đó, và giữ cách gửi này cho các bước sau của lượt.
      const mayThink = opts.provider.kind === "anthropic" || opts.provider.kind === "gemini" || opts.provider.id === "openrouter" || opts.provider.id === "openai";
      if (mayThink && !opts.signal?.aborted && !emitted && opts.state && !opts.state.plain && isThinkingParamError(e)) {
        opts.state.plain = true;
        opts.onThinkingOff?.(String(e?.message || ""));
        attempt--;
        continue;
      }
      if (opts.signal?.aborted || emitted || !isRetryable(e) || attempt >= delays.length) throw e;
      const ra: number | undefined = e.retryAfter;
      if (ra !== undefined && ra > MAX_RETRY_AFTER_S) throw e;
      const waitMs = ra !== undefined ? Math.max(Math.ceil(ra * 1000), delays[attempt]) : delays[attempt];
      const reason =
        e instanceof TypeError
          ? "mất kết nối tới nhà cung cấp"
          : e.status === 429
            ? "hết hạn mức tạm thời"
            : e.status === 529
              ? "nhà cung cấp quá tải"
              : e.status === 408
                ? "nhà cung cấp hết giờ chờ"
                : `máy chủ nhà cung cấp lỗi ${e.status}`;
      opts.onRetry?.({ attempt: attempt + 1, max: delays.length, waitMs, reason });
      await (opts.sleep || sleepAbortable)(waitMs, opts.signal);
    }
  }
}

async function callModelOnce(opts: CallModelOptions): Promise<StepResult> {
  const { url, init } = buildStepRequest(opts.provider, opts.key, opts.model, opts.messages, opts.tools, {
    webSearch: opts.webSearch,
    vision: opts.vision,
    maxOutput: opts.maxOutput,
    think: opts.think,
    plain: opts.state?.plain,
    foreignToolsAsText: opts.foreignToolsAsText,
  });
  const fetchFn = opts.fetchFn || fetch;
  // Chó canh: nhà cung cấp nhận yêu cầu rồi im lặng (đo 16/11/2026 với Gemini lúc quá tải: khung treo vô hạn) thì hủy yêu cầu
  // sau STALL_MS không có byte nào, báo lỗi rõ cho người dùng. Nút Dừng vẫn hủy ngay như cũ.
  const ctl = new AbortController();
  const onUserAbort = () => ctl.abort();
  if (opts.signal?.aborted) ctl.abort();
  opts.signal?.addEventListener("abort", onUserAbort, { once: true });
  let stalled = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const stallMs = opts.stallMs ?? STALL_MS;
  const kick = () => {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      stalled = true;
      ctl.abort();
    }, stallMs);
  };
  const acc = createStepAccumulator(opts.provider.kind, opts.onText, opts.onThinking, opts.provider);
  const parser = createSSEParser((d) => acc.push(d));
  try {
    kick();
    const res = await fetchDiagnosed(opts.provider, url, { ...init, signal: ctl.signal }, fetchFn, opts.model);
    await ensureOk(res, opts.provider);
    if (res.body) {
      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      for (;;) {
        kick();
        const { done, value } = await reader.read();
        if (done) break;
        parser.push(decoder.decode(value, { stream: true }));
      }
    } else {
      parser.push(await res.text());
    }
  } catch (e: any) {
    if (stalled && !opts.signal?.aborted) {
      // ApiError không kèm mã trạng thái: hiện như lỗi của nhà cung cấp, và không thuộc nhóm tự gửi lại (tránh chờ thêm nhiều lần 150 giây)
      const err = new ApiError(`${opts.provider.name}: không nhận được phản hồi trong ${Math.round(stallMs / 1000)} giây (máy chủ nhà cung cấp im lặng). Thử lại, hoặc đổi model / đổi nguồn.`);
      err.stall = true;
      throw err;
    }
    throw e;
  } finally {
    if (timer) clearTimeout(timer);
    opts.signal?.removeEventListener("abort", onUserAbort);
  }
  parser.end();
  return acc.result();
}
