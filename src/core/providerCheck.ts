/** Nút "Kiểm tra nhà cung cấp": chạy thật một chuỗi "gọi công cụ → nhận kết quả → trả lời" với từng key,
 *  và bảng đổi nguồn: lịch sử có gọi công cụ của nguồn A gửi tiếp sang nguồn B, B phải nhắc lại đúng kết quả. */
import { Provider } from "./providers";
import { runAgent } from "./agent/loop";
import { AgentMessage, HostToolSet, ToolSpec } from "./agent/types";
import { describeError } from "./errors";

export const CHECK_TOOL = "lay_ma_kiem_tra";

const SPEC: ToolSpec = {
  name: CHECK_TOOL,
  description: "Trả về mã kiểm tra kỹ thuật gồm 6 chữ số.",
  parameters: { type: "object", properties: {} },
};

const SYSTEM = "Bạn đang được add-in AI Đa Nguồn kiểm tra kỹ thuật. Làm đúng yêu cầu, trả lời thật ngắn bằng tiếng Việt.";

export interface CheckTarget {
  provider: Provider;
  key: string;
  model: string;
  vision?: boolean;
}

export interface CheckResult {
  ok: boolean;
  /** Đạt (n giây) hoặc lý do không đạt, tiếng Việt */
  message: string;
  ms: number;
  /** Lịch sử lượt kiểm (để dùng cho bảng đổi nguồn) */
  history?: AgentMessage[];
}

function testTools(code: string, allowCall: boolean): HostToolSet {
  return {
    host: "Kiểm tra",
    specs: [SPEC],
    overview: async () => ({}),
    systemPrompt: () => SYSTEM,
    execute: async () =>
      allowCall ? { ok: true, ma: code } : { ok: false, error: "Không gọi lại công cụ ở lượt này; hãy dùng mã đã nhận ở lượt trước." },
  };
}

const lastText = (h: AgentMessage[]): string => {
  for (let i = h.length - 1; i >= 0; i--) {
    const m = h[i];
    if (m.role === "assistant" && m.content?.trim()) return m.content;
  }
  return "";
};

const secs = (ms: number) => `${(ms / 1000).toFixed(1).replace(".", ",")} giây`;

export function newCheckCode(): string {
  return String(100000 + Math.floor(Math.random() * 900000));
}

/** Một nguồn: model phải gọi công cụ, nhận mã, rồi trả lời có chứa mã. */
export async function checkProvider(o: CheckTarget & { code: string; fetchFn?: typeof fetch; signal?: AbortSignal }): Promise<CheckResult> {
  const t0 = Date.now();
  let called = false;
  try {
    const history = await runAgent({
      provider: o.provider,
      key: o.key,
      model: o.model,
      vision: o.vision,
      tools: testTools(o.code, true),
      system: SYSTEM,
      messages: [{ role: "user", content: `Gọi công cụ ${CHECK_TOOL} đúng một lần, rồi trả lời đúng một câu: "Mã là <mã nhận được>".` }],
      maxSteps: 3,
      signal: o.signal,
      log: [],
      fetchFn: o.fetchFn,
      onEvent: (ev) => {
        if (ev.type === "tool_start") called = true;
      },
    });
    const ms = Date.now() - t0;
    const text = lastText(history);
    if (!called) return { ok: false, ms, message: "Model không gọi công cụ, nên không thao tác được trên tài liệu. Hãy chọn model khác." };
    if (!text.includes(o.code)) return { ok: false, ms, history, message: `Có gọi công cụ nhưng câu trả lời không chứa kết quả (trả lời: “${text.slice(0, 100)}”).` };
    return { ok: true, ms, history, message: `Đạt (${secs(ms)})` };
  } catch (e) {
    return { ok: false, ms: Date.now() - t0, message: describeError(e, { provider: o.provider }).text };
  }
}

/** Đổi nguồn giữa hội thoại: gửi lịch sử của nguồn A cho nguồn B, B phải nhắc lại đúng mã đã nhận ở lượt trước. */
export async function checkSwitch(o: {
  history: AgentMessage[];
  code: string;
  to: CheckTarget;
  fetchFn?: typeof fetch;
  signal?: AbortSignal;
}): Promise<CheckResult> {
  const t0 = Date.now();
  try {
    const history = await runAgent({
      provider: o.to.provider,
      key: o.to.key,
      model: o.to.model,
      vision: o.to.vision,
      tools: testTools(o.code, false),
      system: SYSTEM,
      messages: [...o.history, { role: "user", content: "Nhắc lại mã kiểm tra đã nhận ở lượt trước. Chỉ ghi các chữ số, không gọi công cụ." }],
      maxSteps: 3,
      signal: o.signal,
      log: [],
      fetchFn: o.fetchFn,
      onEvent: () => {},
    });
    const ms = Date.now() - t0;
    const text = lastText(history);
    if (!text.includes(o.code)) return { ok: false, ms, message: `Chạy được nhưng không nhớ lượt trước (trả lời: “${text.slice(0, 80)}”).` };
    return { ok: true, ms, message: `Đạt (${secs(ms)})` };
  } catch (e) {
    return { ok: false, ms: Date.now() - t0, message: describeError(e, { provider: o.to.provider }).text };
  }
}

export interface MatrixCell {
  from: string;
  to: string;
  ok: boolean;
  message: string;
}

const wait = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener("abort", () => {
      clearTimeout(t);
      reject(new DOMException("Đã dừng", "AbortError"));
    });
  });

/** Gặp 429 (hết hạn mức tạm thời) thì chờ rồi thử lại một lần. */
async function withRetry(fn: () => Promise<CheckResult>, signal?: AbortSignal): Promise<CheckResult> {
  const r = await fn();
  if (r.ok || !/Hết hạn mức/.test(r.message)) return r;
  await wait(20000, signal);
  return fn();
}

/** Chạy từng nguồn rồi mọi cặp đổi nguồn (cả hai chiều). Thứ tự xen kẽ để không dồn nhiều lượt vào một nguồn liền nhau. */
export async function runSwitchMatrix(o: {
  targets: CheckTarget[];
  onSingle: (providerId: string, r: CheckResult) => void;
  onCell: (cell: MatrixCell) => void;
  signal?: AbortSignal;
  fetchFn?: typeof fetch;
  /** Chỉ kiểm từng nguồn, không chạy bảng đổi nguồn */
  singlesOnly?: boolean;
}): Promise<{ singles: Record<string, CheckResult>; cells: MatrixCell[] }> {
  const singles: Record<string, CheckResult> = {};
  const codes: Record<string, string> = {};
  for (const t of o.targets) {
    if (o.signal?.aborted) throw new DOMException("Đã dừng", "AbortError");
    codes[t.provider.id] = newCheckCode();
    const r = await withRetry(() => checkProvider({ ...t, code: codes[t.provider.id], fetchFn: o.fetchFn, signal: o.signal }), o.signal);
    singles[t.provider.id] = r;
    o.onSingle(t.provider.id, r);
  }
  const cells: MatrixCell[] = [];
  if (o.singlesOnly) return { singles, cells };
  const n = o.targets.length;
  for (let d = 1; d < n; d++) {
    for (let i = 0; i < n; i++) {
      if (o.signal?.aborted) throw new DOMException("Đã dừng", "AbortError");
      const a = o.targets[i];
      const b = o.targets[(i + d) % n];
      const ra = singles[a.provider.id];
      let cell: MatrixCell;
      if (!ra.ok || !ra.history) cell = { from: a.provider.id, to: b.provider.id, ok: false, message: `Bỏ qua: ${a.provider.name} chưa đạt khi kiểm riêng` };
      else if (!singles[b.provider.id].ok && !/không gọi công cụ|không chứa kết quả/.test(singles[b.provider.id].message))
        cell = { from: a.provider.id, to: b.provider.id, ok: false, message: `Bỏ qua: ${b.provider.name} chưa đạt khi kiểm riêng` };
      else {
        const r = await withRetry(
          () => checkSwitch({ history: ra.history!, code: codes[a.provider.id], to: b, fetchFn: o.fetchFn, signal: o.signal }),
          o.signal
        );
        cell = { from: a.provider.id, to: b.provider.id, ok: r.ok, message: r.message };
      }
      cells.push(cell);
      o.onCell(cell);
    }
  }
  return { singles, cells };
}

/** Bảng kết quả dạng chữ để chép làm bằng chứng. */
export function matrixReport(targets: CheckTarget[], singles: Record<string, CheckResult>, cells: MatrixCell[], when = new Date()): string {
  const lines = [`Kiểm tra nhà cung cấp — ${when.toLocaleString("vi-VN")}`, "", "Từng nguồn (gọi công cụ → nhận kết quả → trả lời):"];
  for (const t of targets) {
    const r = singles[t.provider.id];
    lines.push(`- ${t.provider.name} · ${t.model}: ${r ? (r.ok ? "ĐẠT" : "KHÔNG ĐẠT") + " — " + r.message : "chưa chạy"}`);
  }
  if (cells.length) {
    const ok = cells.filter((c) => c.ok).length;
    lines.push("", `Đổi nguồn giữa hội thoại (A → B, B nhắc lại kết quả công cụ của A): đạt ${ok}/${cells.length}`);
    for (const c of cells) if (!c.ok) lines.push(`- ${c.from} → ${c.to}: ${c.message}`);
  }
  return lines.join("\n");
}

