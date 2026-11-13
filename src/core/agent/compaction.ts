/** Tự thu gọn hội thoại dài: tóm tắt phần cũ bằng chính model đang dùng, giữ nguyên các lượt gần nhất. */
import { Provider } from "../providers";
import { callModelStep } from "./llm";
import { AgentMessage } from "./types";

export const COMPACT_AT_CHARS = 120000;
export const KEEP_RECENT_TURNS = 2;

export function historyChars(h: AgentMessage[]): number {
  return h.reduce((n, m) => n + (m.content?.length || 0) + JSON.stringify((m as any).toolCalls || "").length, 0);
}

/** Cắt tại đầu một lượt người dùng để không làm vỡ cặp gọi công cụ – kết quả. */
export function splitForCompaction(h: AgentMessage[], keepTurns = KEEP_RECENT_TURNS): { old: AgentMessage[]; recent: AgentMessage[] } {
  const userIdx = h.map((m, i) => (m.role === "user" ? i : -1)).filter((i) => i >= 0);
  if (userIdx.length <= keepTurns) return { old: [], recent: h };
  const cut = userIdx[userIdx.length - keepTurns];
  return { old: h.slice(0, cut), recent: h.slice(cut) };
}

export function transcriptOf(h: AgentMessage[], maxChars = 60000): string {
  const lines: string[] = [];
  for (const m of h) {
    if (m.role === "user") lines.push(`NGƯỜI DÙNG: ${m.content}`);
    else if (m.role === "assistant") {
      if (m.content) lines.push(`TRỢ LÝ: ${m.content}`);
      for (const tc of m.toolCalls || []) lines.push(`TRỢ LÝ gọi ${tc.name}(${JSON.stringify(tc.args).slice(0, 300)})`);
    } else if (m.role === "tool") lines.push(`KẾT QUẢ ${m.name}: ${m.content.slice(0, 600)}`);
  }
  const t = lines.join("\n");
  return t.length > maxChars ? t.slice(t.length - maxChars) : t;
}

export async function compactIfNeeded(o: {
  history: AgentMessage[];
  provider: Provider;
  key: string;
  model: string;
  limitChars?: number;
  signal?: AbortSignal;
  fetchFn?: typeof fetch;
}): Promise<{ history: AgentMessage[]; compacted: boolean }> {
  if (historyChars(o.history) < (o.limitChars ?? COMPACT_AT_CHARS)) return { history: o.history, compacted: false };
  const { old, recent } = splitForCompaction(o.history);
  if (!old.length) return { history: o.history, compacted: false };
  const res = await callModelStep({
    provider: o.provider,
    key: o.key,
    model: o.model,
    tools: [],
    signal: o.signal,
    onText: () => {},
    fetchFn: o.fetchFn,
    messages: [
      {
        role: "system",
        content:
          "Tóm tắt đoạn hội thoại dưới đây thành bản ghi nhớ ngắn gọn bằng tiếng Việt để làm tiếp công việc: mục tiêu của người dùng, những gì đã làm trong tài liệu (ô, đoạn, slide cụ thể), số liệu/quyết định quan trọng, việc còn dở. Không bịa thêm.",
      },
      { role: "user", content: transcriptOf(old) },
    ],
  });
  const summary = res.text.trim() || "(không tóm tắt được)";
  return {
    compacted: true,
    history: [
      { role: "user", content: `[Tóm tắt phần hội thoại trước đã được thu gọn]\n${summary}` },
      { role: "assistant", content: "Đã nắm nội dung trước đó, tôi sẽ làm tiếp." },
      ...recent,
    ],
  };
}
