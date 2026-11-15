/** Điều phối một lượt hội thoại (tách khỏi giao diện để kiểm thử được ngoài Office):
 *  tóm tắt khi dài → dựng lời nhắc và tin nhắn của lượt → chạy tác tử hoặc trò chuyện thường → trả lịch sử mới.
 *  Không ném lỗi: bị dừng hay lỗi thì trả `error` kèm lịch sử cần giữ (phần đã làm vẫn còn; chưa làm gì thì như chưa gửi). */
import { Provider } from "../providers";
import { compactIfNeeded } from "./compaction";
import { callModelStep } from "./llm";
import { STOP_NOTE, runAgent } from "./loop";
import { buildPlainSystem, buildSystemPrompt, buildTurnMessage } from "./prompt";
import { AgentEvent, AgentMessage, ExtraTools, HostToolSet, ToolCall, UndoEntry, Usage } from "./types";

export type TurnEvent = AgentEvent | { type: "compacted" };

export interface TurnAgentOptions {
  tools: HostToolSet;
  extra?: ExtraTools;
  /** Chuẩn làm việc theo ứng dụng */
  guide?: string;
  skillCatalog?: string;
  /** Kỹ năng tự chọn cho yêu cầu này (đưa vào tin nhắn của lượt) */
  autoSkills?: { name: string; body: string }[];
  askBeforeWrite?: boolean;
  confirmRisky?: boolean;
  approve?: (call: ToolCall, reason?: string) => Promise<boolean>;
  /** Nhật ký hoàn tác của lượt */
  log: UndoEntry[];
  maxSteps?: number;
  toolTimeoutMs?: number;
}

export interface TurnOptions {
  provider: Provider;
  key: string;
  model: string;
  /** Thông tin model nếu biết: sức chứa, trần đầu ra, nhận ảnh */
  context?: number;
  maxOutput?: number;
  vision?: boolean;
  host: string;
  /** Lịch sử trước lượt này (không bị sửa) */
  history: AgentMessage[];
  /** Lời người dùng, đã ghép vùng chọn và chữ của tệp đính kèm */
  text: string;
  images?: { mime: string; data: string }[];
  /** Chỉ dẫn riêng của người dùng cho ứng dụng */
  instructions?: string;
  /** Có = chế độ tác tử; không = trò chuyện thường */
  agent?: TurnAgentOptions;
  webSearch?: boolean;
  /** Token đầu vào thực đo của bước gần nhất trong hội thoại này */
  lastInputTokens?: number;
  signal?: AbortSignal;
  onEvent: (e: TurnEvent) => void;
  /** Lịch sử vừa có thêm tin nhắn (để lưu theo từng bước) */
  onHistory?: (h: AgentMessage[]) => void;
  fetchFn?: typeof fetch;
  retryDelaysMs?: number[];
  sleep?: (ms: number, signal?: AbortSignal) => Promise<void>;
}

export interface TurnResult {
  /** Lịch sử cần giữ sau lượt */
  history: AgentMessage[];
  compacted: boolean;
  /** Tổng token của lượt (mọi bước) */
  usage: Usage;
  /** Token đầu vào của bước cuối — đo độ đầy của ngữ cảnh cho lượt sau */
  lastInputTokens?: number;
  error?: unknown;
}

export function addUsage(a: Usage, b?: Usage): Usage {
  if (!b) return a;
  const sum = (x?: number, y?: number) => (x === undefined && y === undefined ? undefined : (x || 0) + (y || 0));
  return { input: sum(a.input, b.input), output: sum(a.output, b.output), cacheRead: sum(a.cacheRead, b.cacheRead), cacheWrite: sum(a.cacheWrite, b.cacheWrite) };
}

/** Dòng hiện số token của một lượt, ví dụ "Token: vào 12.345 · ra 678 · đọc từ bộ nhớ đệm 11.000". Không có số → chuỗi rỗng. */
export function formatUsage(u?: Usage): string {
  if (!u || (!u.input && !u.output)) return "";
  const n = (v: number) => Math.round(v).toLocaleString("vi-VN");
  const parts = [`vào ${n(u.input || 0)}`, `ra ${n(u.output || 0)}`];
  if (u.cacheRead) parts.push(`đọc từ bộ nhớ đệm ${n(u.cacheRead)}`);
  if (u.cacheWrite) parts.push(`ghi vào bộ nhớ đệm ${n(u.cacheWrite)}`);
  return `Token: ${parts.join(" · ")}`;
}

/** Hội thoại gửi model ở chế độ trò chuyện thường: bỏ lời gọi công cụ */
export const plainHistory = (h: AgentMessage[]): AgentMessage[] =>
  h
    .filter((m) => m.role === "user" || (m.role === "assistant" && m.content))
    .map((m) => (m.role === "assistant" ? { role: "assistant", content: m.content } : { role: "user", content: m.content })) as AgentMessage[];

export async function runTurn(o: TurnOptions): Promise<TurnResult> {
  let history = o.history;
  let compacted = false;
  let usage: Usage = {};
  let lastInputTokens = o.lastInputTokens;
  const result = (h: AgentMessage[], error?: unknown): TurnResult => ({ history: h, compacted, usage, lastInputTokens, error });
  const onUsage = (u?: Usage) => {
    if (!u) return;
    usage = addUsage(usage, u);
    if (u.input !== undefined) lastInputTokens = u.input;
  };

  try {
    const comp = await compactIfNeeded({
      history,
      provider: o.provider,
      key: o.key,
      model: o.model,
      context: o.context,
      lastInputTokens: o.lastInputTokens,
      signal: o.signal,
      fetchFn: o.fetchFn,
    });
    if (comp.compacted) {
      history = comp.history;
      compacted = true;
      lastInputTokens = undefined;
      o.onEvent({ type: "compacted" });
      o.onHistory?.(history);
    }
  } catch (e) {
    return result(history, e);
  }

  const userOf = (content: string): AgentMessage => (o.images?.length ? { role: "user", content, images: o.images } : { role: "user", content });

  if (o.agent) {
    const a = o.agent;
    let ov: any;
    try {
      ov = await a.tools.overview();
    } catch (e: any) {
      ov = { error: "Không đọc được trạng thái tài liệu: " + (e?.message || String(e)) };
    }
    const system = buildSystemPrompt({
      base: a.tools.systemPrompt(),
      guide: a.guide,
      host: o.host,
      instructions: o.instructions,
      skillCatalog: a.skillCatalog,
    });
    const userMsg = userOf(buildTurnMessage({ text: o.text, state: a.tools.stateText?.(ov), skills: a.autoSkills }));
    const start = [...history, userMsg];
    let latest = start;
    try {
      const msgs = await runAgent({
        provider: o.provider,
        key: o.key,
        model: o.model,
        tools: a.tools,
        extra: a.extra,
        system,
        messages: start,
        signal: o.signal,
        log: a.log,
        approve: a.approve,
        askBeforeWrite: a.askBeforeWrite,
        confirmRisky: a.confirmRisky,
        webSearch: o.webSearch,
        vision: o.vision,
        maxOutput: o.maxOutput,
        maxSteps: a.maxSteps,
        toolTimeoutMs: a.toolTimeoutMs,
        fetchFn: o.fetchFn,
        retryDelaysMs: o.retryDelaysMs,
        sleep: o.sleep,
        onHistory: (h) => {
          latest = h;
          o.onHistory?.(h);
        },
        onEvent: (ev) => {
          if (ev.type === "usage") onUsage(ev.usage);
          o.onEvent(ev);
        },
      });
      return result(msgs);
    } catch (e) {
      // Đã có phần làm được (trả lời dở, bước công cụ) → giữ; chưa có gì → coi như chưa gửi câu hỏi này
      return result(latest.length > start.length ? latest : history, e);
    }
  }

  const userMsg = userOf(o.text);
  let partial = "";
  try {
    const res = await callModelStep({
      provider: o.provider,
      key: o.key,
      model: o.model,
      tools: [],
      messages: [{ role: "system", content: buildPlainSystem(o.instructions) }, ...plainHistory(history), userMsg],
      signal: o.signal,
      onText: (d) => {
        partial += d;
        o.onEvent({ type: "text", delta: d });
      },
      onRetry: (r) => o.onEvent({ type: "retry", ...r }),
      webSearch: o.webSearch,
      vision: o.vision,
      maxOutput: o.maxOutput,
      fetchFn: o.fetchFn,
      retryDelaysMs: o.retryDelaysMs,
      sleep: o.sleep,
    });
    if (res.usage) {
      onUsage(res.usage);
      o.onEvent({ type: "usage", step: 1, usage: res.usage });
    }
    if (res.truncated) o.onEvent({ type: "truncated", skipped: 0 });
    const next: AgentMessage[] = [...history, userMsg, { role: "assistant", content: res.text }];
    o.onHistory?.(next);
    return result(next);
  } catch (e: any) {
    if (!partial.trim()) return result(history, e);
    const stopped = o.signal?.aborted || e?.name === "AbortError";
    const note = stopped ? STOP_NOTE : "[Câu trả lời bị ngắt giữa chừng do lỗi khi gọi nhà cung cấp.]";
    const next: AgentMessage[] = [...history, userMsg, { role: "assistant", content: `${partial}\n\n${note}` }];
    o.onHistory?.(next);
    return result(next, e);
  }
}
