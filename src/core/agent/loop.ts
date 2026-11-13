/** Vòng lặp tác tử: model đề xuất lệnh gọi công cụ → add-in thực hiện trên tài liệu → trả kết quả → lặp đến khi xong. */
import { Provider } from "../providers";
import { callModelStep } from "./llm";
import { AgentEvent, AgentMessage, HostToolSet, ToolCall, UndoEntry } from "./types";

export const MAX_TOOL_RESULT_CHARS = 30000;
export const OLD_TOOL_RESULT_CHARS = 1500;

export function clip(s: string, max: number): string {
  return s.length <= max ? s : s.slice(0, max) + `… [đã cắt bớt ${s.length - max} ký tự]`;
}

/** Thu gọn kết quả công cụ của các lượt cũ để tiết kiệm ngữ cảnh. */
export function compactHistory(history: AgentMessage[]): AgentMessage[] {
  return history.map((m) => (m.role === "tool" ? { ...m, content: clip(m.content, OLD_TOOL_RESULT_CHARS) } : m));
}

export interface RunAgentOptions {
  provider: Provider;
  key: string;
  model: string;
  tools: HostToolSet;
  /** Hội thoại trước đó (không gồm system) + câu hỏi mới ở cuối */
  messages: AgentMessage[];
  system: string;
  maxSteps?: number;
  signal?: AbortSignal;
  log: UndoEntry[];
  onEvent: (e: AgentEvent) => void;
  /** Trả về false nếu người dùng không cho phép thao tác ghi */
  approve?: (call: ToolCall) => Promise<boolean>;
  fetchFn?: typeof fetch;
}

/** Trả về toàn bộ hội thoại sau lần chạy (không gồm system). */
export async function runAgent(o: RunAgentOptions): Promise<AgentMessage[]> {
  const maxSteps = o.maxSteps ?? 20;
  const msgs: AgentMessage[] = [...o.messages];
  const writeTools = new Set(o.tools.specs.filter((t) => t.write).map((t) => t.name));
  const known = new Set(o.tools.specs.map((t) => t.name));

  for (let step = 1; step <= maxSteps; step++) {
    if (o.signal?.aborted) throw new DOMException("Đã dừng", "AbortError");
    o.onEvent({ type: "step", step });
    const res = await callModelStep({
      provider: o.provider,
      key: o.key,
      model: o.model,
      messages: [{ role: "system", content: o.system }, ...msgs],
      tools: o.tools.specs,
      signal: o.signal,
      onText: (d) => o.onEvent({ type: "text", delta: d }),
      fetchFn: o.fetchFn,
    });
    msgs.push({ role: "assistant", content: res.text, toolCalls: res.toolCalls, raw: res.raw });
    if (!res.toolCalls.length) return msgs;

    for (const call of res.toolCalls) {
      if (o.signal?.aborted) throw new DOMException("Đã dừng", "AbortError");
      o.onEvent({ type: "tool_start", call });
      let result: any;
      let error: string | undefined;
      try {
        if (!known.has(call.name)) throw new Error(`Không có công cụ tên "${call.name}".`);
        if (call.args?.__invalid_json) throw new Error("Tham số gửi lên không phải JSON hợp lệ, hãy gọi lại.");
        if (writeTools.has(call.name) && o.approve && !(await o.approve(call))) {
          throw new Error("Người dùng không cho phép thao tác này. Đừng thử lại; hỏi người dùng muốn làm gì khác.");
        }
        result = await o.tools.execute(call.name, call.args || {}, { log: o.log });
      } catch (e: any) {
        error = e?.message || String(e);
        if (e?.code && !String(error).includes(e.code)) error += ` [${e.code}]`;
        if (e?.debugInfo?.errorLocation) error += ` (${e.debugInfo.errorLocation})`;
        result = { ok: false, error };
      }
      o.onEvent({ type: "tool_end", call, result, error });
      msgs.push({
        role: "tool",
        toolCallId: call.id,
        name: call.name,
        content: clip(JSON.stringify(result ?? { ok: true }), MAX_TOOL_RESULT_CHARS),
      });
    }
  }
  o.onEvent({ type: "limit", maxSteps });
  return msgs;
}
