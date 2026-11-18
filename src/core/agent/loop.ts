/** Vòng lặp tác tử: model đề xuất lệnh gọi công cụ → add-in thực hiện trên tài liệu → trả kết quả → lặp đến khi xong.
 *  Lịch sử CHỈ THÊM: mỗi tin nhắn được ghi ngay khi có (qua onHistory), kể cả khi người dùng bấm Dừng hoặc gặp lỗi,
 *  và mọi lệnh gọi công cụ đều có kết quả đi kèm — để lượt sau AI biết chính xác mình đã làm gì. */
import { Provider } from "../providers";
import { callModelStep } from "./llm";
import { AgentEvent, AgentMessage, ExtraTools, HostToolSet, StepResult, ToolCall, ToolSpec, UndoEntry } from "./types";

export const MAX_TOOL_RESULT_CHARS = 30000;
/** Hạn chờ mặc định cho một bước công cụ (bản chính thức: 5 phút mỗi bước, 30 giây cho một sửa chữ ở Word). */
export const TOOL_TIMEOUT_MS = 120000;
/** Câu trả lời bị cắt vì quá dài mấy lần liên tiếp thì dừng lượt, không để model lặp vô ích. */
export const MAX_TRUNCATED_IN_A_ROW = 2;

export const STOP_NOTE = "[Người dùng đã bấm Dừng ở đây. Các bước ở trên có kết quả thành công thì đã được thực hiện trong tài liệu.]";
export const NOT_RUN_STOPPED = "Chưa thực hiện: người dùng đã bấm Dừng trước bước này.";
/** Bước soát lập luận chặn tối đa bấy nhiêu lần trong một lượt; sau đó lệnh ghi được chạy thẳng (tránh vòng lặp và tốn token) */
export const MAX_REVIEW_BLOCKS = 2;
export const NOT_RUN_TRUNCATED =
  "Chưa thực hiện: câu trả lời vừa rồi bị cắt vì vượt giới hạn độ dài nên lệnh gọi công cụ có thể dở dang. Hãy gọi lại, chia nhỏ: mỗi lần ghi ít nội dung hơn (ví dụ từng phần của văn bản, từng nhóm hàng, từng slide).";

/** Cùng một công cụ báo cùng một lỗi mấy bước liên tiếp thì dừng lượt, không đợi hết trần số bước:
 *  lặp y nguyên tham số → 2 bước; tham số có đổi mà lỗi vẫn thế → 3 bước. */
export const STUCK_SAME_ARGS = 2;
export const STUCK_SAME_ERROR = 3;

/** Trần số bước của một lượt theo ứng dụng (số đo Chạy thật: Word và Excel 3–4 lần gọi AI cho một việc, PowerPoint 6–8).
 *  Chạm trần thì khung hỏi "Làm tiếp?" chứ không tự chạy tiếp. */
export function stepCapFor(host: string): number {
  return host === "PowerPoint" ? 20 : host === "Word" || host === "Excel" ? 12 : 20;
}

export function clip(s: string, max: number): string {
  return s.length <= max ? s : s.slice(0, max) + `… [đã cắt bớt ${s.length - max} ký tự]`;
}

/** Chờ một bước công cụ có hạn. Hết hạn thì thôi chờ (không hủy được thao tác Office đang chạy) và báo cho model. */
export function withTimeout<T>(p: Promise<T>, ms: number, what: string): Promise<T> {
  if (!Number.isFinite(ms) || ms <= 0) return p;
  return new Promise<T>((resolve, reject) => {
    const t = setTimeout(
      () =>
        reject(
          new Error(
            `Bước "${what}" chạy quá ${Math.round(ms / 1000)} giây nên add-in thôi chờ. Thao tác có thể vẫn đang chạy hoặc đã xong trong tài liệu: hãy đọc lại trạng thái tài liệu trước khi làm tiếp, không lặp lại ngay thao tác ghi.`
          )
        ),
      ms
    );
    p.then(
      (v) => {
        clearTimeout(t);
        resolve(v);
      },
      (e) => {
        clearTimeout(t);
        reject(e);
      }
    );
  });
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
  /** Được gọi mỗi khi lịch sử có thêm tin nhắn (cùng một mảng, chỉ thêm) — để lưu theo từng bước */
  onHistory?: (messages: AgentMessage[]) => void;
  /** Hỏi người dùng; trả về false nếu không cho phép. reason = cảnh báo rủi ro (nếu có) */
  approve?: (call: ToolCall, reason?: string) => Promise<boolean>;
  /** Lời người dùng của lượt này (chưa ghép phần add-in tự thêm) — chỉ để công cụ ghi nhật ký soát lại, không gửi model */
  request?: string;
  /** Hỏi trước mọi thao tác ghi */
  askBeforeWrite?: boolean;
  /** Luôn hỏi trước thao tác rủi ro (ghi đè dữ liệu, xóa) */
  confirmRisky?: boolean;
  /** Công cụ dùng chung (use_skill, read_other_app…) */
  extra?: ExtraTools;
  webSearch?: boolean;
  /** Model nhận ảnh (true/false/không rõ) — để bỏ ảnh cũ khi đổi sang model không nhận ảnh */
  vision?: boolean;
  /** Trần độ dài đầu ra của model nếu biết */
  maxOutput?: number;
  /** Hạn chờ cho mỗi bước công cụ (mili giây); 0 = không giới hạn */
  toolTimeoutMs?: number;
  /** Công tắc "Suy nghĩ kỹ" */
  think?: boolean;
  /** Soát lập luận của một lệnh ghi vào điều khoản trước khi chạy (công cụ có reviewable). ok = false thì lệnh không chạy,
   *  note được trả cho model để sửa đề xuất. Lỗi khi soát không chặn lệnh. */
  review?: (call: ToolCall, spec: ToolSpec) => Promise<{ ok: boolean; note?: string }>;
  fetchFn?: typeof fetch;
  /** Cho kiểm thử: bảng khoảng chờ thử lại và hàm chờ */
  retryDelaysMs?: number[];
  sleep?: (ms: number, signal?: AbortSignal) => Promise<void>;
}

const stopError = () => new DOMException("Đã dừng", "AbortError");

/** Trả về toàn bộ hội thoại sau lần chạy (không gồm system). Khi bị dừng hoặc lỗi, phần đã làm vẫn nằm trong
 *  mảng đã báo qua onHistory. */
export async function runAgent(o: RunAgentOptions): Promise<AgentMessage[]> {
  const maxSteps = o.maxSteps ?? 20;
  const msgs: AgentMessage[] = [...o.messages];
  const base = msgs.length;
  const allSpecs: ToolSpec[] = [...o.tools.specs, ...(o.extra?.specs || [])];
  const specMap = new Map(allSpecs.map((t) => [t.name, t]));
  const extraNames = new Set((o.extra?.specs || []).map((t) => t.name));

  const push = (m: AgentMessage) => {
    msgs.push(m);
    o.onHistory?.(msgs);
  };
  const toolMsg = (call: ToolCall, result: any): AgentMessage => ({
    role: "tool",
    toolCallId: call.id,
    name: call.name,
    content: clip(JSON.stringify(result ?? { ok: true }), MAX_TOOL_RESULT_CHARS),
  });
  const aborted = () => !!o.signal?.aborted;
  /** Ghi mốc dừng. Chưa làm được gì trong lượt này thì không ghi (bên gọi sẽ bỏ luôn câu hỏi của lượt). */
  const markStopped = (partial = "") => {
    if (partial.trim()) push({ role: "assistant", content: `${partial}\n\n${STOP_NOTE}` });
    else if (msgs.length > base) push({ role: "assistant", content: STOP_NOTE });
  };

  let truncatedRun = 0;
  let reviewBlocks = 0;
  /** Dùng chung cho mọi bước của lượt: nhà cung cấp đã từ chối tham số suy nghĩ thì các bước sau không gửi nữa */
  const callState: { plain?: boolean } = {};
  /** Số lệnh ghi đã chạy thành công trong lượt (để thẻ "Làm tiếp?" nêu rõ đã làm gì) */
  let writes = 0;
  /** Lỗi công cụ của bước trước: "tên|lỗi" → { số bước liên tiếp, số bước liên tiếp y nguyên tham số, tham số lần cuối } */
  let prevFails = new Map<string, { run: number; sameArgs: number; args: string }>();
  for (let step = 1; step <= maxSteps; step++) {
    if (aborted()) {
      markStopped();
      throw stopError();
    }
    o.onEvent({ type: "step", step });
    let partial = "";
    let res: StepResult;
    try {
      res = await callModelStep({
        provider: o.provider,
        key: o.key,
        model: o.model,
        messages: [{ role: "system", content: o.system }, ...msgs],
        tools: allSpecs,
        signal: o.signal,
        onText: (d) => {
          partial += d;
          o.onEvent({ type: "text", delta: d });
        },
        onRetry: (r) => o.onEvent({ type: "retry", ...r }),
        onThinking: () => o.onEvent({ type: "thinking" }),
        onThinkingOff: (reason) => o.onEvent({ type: "thinking_off", reason }),
        think: o.think,
        state: callState,
        fetchFn: o.fetchFn,
        webSearch: o.webSearch,
        vision: o.vision,
        maxOutput: o.maxOutput,
        retryDelaysMs: o.retryDelaysMs,
        sleep: o.sleep,
      });
    } catch (e: any) {
      if (aborted() || e?.name === "AbortError") markStopped(partial);
      else if (partial.trim()) push({ role: "assistant", content: `${partial}\n\n[Câu trả lời bị ngắt giữa chừng do lỗi khi gọi nhà cung cấp.]` });
      else if (msgs.length > base) push({ role: "assistant", content: "[Lượt này dừng ở đây do lỗi khi gọi nhà cung cấp. Các bước ở trên có kết quả thành công thì đã được thực hiện.]" });
      throw e;
    }
    if (res.usage) o.onEvent({ type: "usage", step, usage: res.usage });
    push({ role: "assistant", content: res.text, toolCalls: res.toolCalls, raw: res.raw });

    if (res.truncated) {
      // Hết độ dài đầu ra: lệnh gọi công cụ có thể bị cắt dở → không chạy lệnh nào của bước này
      truncatedRun++;
      for (const call of res.toolCalls) {
        o.onEvent({ type: "tool_start", call });
        o.onEvent({ type: "tool_end", call, result: { ok: false, error: NOT_RUN_TRUNCATED }, error: "Chưa chạy: lệnh bị cắt dở vì câu trả lời quá dài" });
        push(toolMsg(call, { ok: false, error: NOT_RUN_TRUNCATED }));
      }
      const gaveUp = res.toolCalls.length > 0 && truncatedRun >= MAX_TRUNCATED_IN_A_ROW;
      o.onEvent({ type: "truncated", skipped: res.toolCalls.length, gaveUp });
      if (!res.toolCalls.length) return msgs;
      if (gaveUp) {
        push({ role: "assistant", content: "[Dừng lượt: câu trả lời bị cắt vì quá dài nhiều lần liên tiếp, các lệnh ghi dở chưa được thực hiện.]" });
        return msgs;
      }
      continue;
    }
    truncatedRun = 0;
    if (!res.toolCalls.length) return msgs;

    /** Ảnh do công cụ trả về trong bước này (ảnh slide…): gửi cho model ở một tin nhắn riêng ngay sau các kết quả công cụ */
    const toolImages: { mime: string; data: string; label?: string }[] = [];
    const stepFails = new Map<string, { run: number; sameArgs: number; args: string }>();
    let stuck: { tool: string; error: string; times: number } | undefined;
    for (let i = 0; i < res.toolCalls.length; i++) {
      const call = res.toolCalls[i];
      if (aborted()) {
        // Mọi lệnh gọi đều phải có kết quả: các lệnh chưa chạy ghi rõ là chưa thực hiện
        for (const rest of res.toolCalls.slice(i)) push(toolMsg(rest, { ok: false, error: NOT_RUN_STOPPED }));
        markStopped();
        throw stopError();
      }
      // Soát lập luận trước khi ghi sửa đổi vào điều khoản (tùy chọn): một lượt model riêng, không nằm trong lịch sử.
      // Chạy trước khi hiện bước ghi và trước khi hỏi người dùng.
      let blocked = "";
      const reviewSpec = specMap.get(call.name);
      if (o.review && reviewSpec?.reviewable && !call.args?.__invalid_json && reviewBlocks < MAX_REVIEW_BLOCKS) {
        o.onEvent({ type: "review_start", call });
        let verdict: { ok: boolean; note?: string };
        try {
          verdict = await o.review(call, reviewSpec);
        } catch (e: any) {
          if (aborted() || e?.name === "AbortError") {
            // Bấm Dừng trong lúc soát: lệnh này và các lệnh sau đều chưa chạy
            o.onEvent({ type: "review_end", call, ok: false, note: "Đã dừng" });
            for (const rest of res.toolCalls.slice(i)) push(toolMsg(rest, { ok: false, error: NOT_RUN_STOPPED }));
            markStopped();
            throw stopError();
          }
          verdict = { ok: true, note: "Không soát được: " + (e?.message || String(e)) };
        }
        o.onEvent({ type: "review_end", call, ok: verdict.ok, note: verdict.note });
        if (!verdict.ok) {
          reviewBlocks++;
          blocked = `CHƯA GHI vào văn bản. Bước soát lập luận trước khi sửa nêu vấn đề: ${verdict.note || "(không nêu lý do)"} — Hãy xem lại đề xuất: nếu vấn đề đúng thì sửa đề xuất rồi gọi lại công cụ; nếu bạn chắc đề xuất ban đầu đúng thì nói rõ lý do với người dùng và hỏi ý họ trước khi ghi.`;
        }
      }
      o.onEvent({ type: "tool_start", call });
      let result: any;
      let error: string | undefined;
      try {
        const spec = specMap.get(call.name);
        if (!spec) throw new Error(`Không có công cụ tên "${call.name}".`);
        if (call.args?.__invalid_json) throw new Error("Tham số gửi lên không phải JSON hợp lệ, hãy gọi lại.");
        if (blocked) throw new Error(blocked);
        let reason: string | null = null;
        if ((o.confirmRisky || spec.alwaysAsk) && spec.risky) reason = await spec.risky(call.args || {});
        if (reason || (spec.write && o.askBeforeWrite)) {
          if (!o.approve || !(await o.approve(call, reason || undefined))) {
            throw new Error("Người dùng không cho phép thao tác này. Đừng thử lại; hỏi người dùng muốn làm gì khác.");
          }
        }
        // Hạn chờ chỉ tính lúc công cụ chạy, không tính lúc chờ người dùng duyệt
        const run = extraNames.has(call.name)
          ? o.extra!.execute(call.name, call.args || {})
          : o.tools.execute(call.name, call.args || {}, { log: o.log, said: res.text, request: o.request });
        result = await withTimeout(Promise.resolve(run), spec.timeoutMs ?? o.toolTimeoutMs ?? TOOL_TIMEOUT_MS, call.name);
      } catch (e: any) {
        error = e?.message || String(e);
        if (e?.code && !String(error).includes(e.code)) error += ` [${e.code}]`;
        if (e?.debugInfo?.errorLocation) error += ` (${e.debugInfo.errorLocation})`;
        result = { ok: false, error };
      }
      // Ảnh không nằm trong kết quả dạng chữ: tách ra, gửi kèm ở tin nhắn sau (model không nhận ảnh thì báo rõ, không gửi)
      if (result && typeof result === "object" && Array.isArray(result.__images)) {
        const imgs = result.__images.filter((x: any) => x && typeof x.data === "string" && typeof x.mime === "string").slice(0, 4);
        const { __images: _drop, ...rest } = result;
        void _drop;
        result = rest;
        if (o.vision === false) result = { ...result, image_note: "Model đang dùng không nhận ảnh nên ảnh không được gửi: chỉ dùng được phần chữ và số đo bằng mã. Muốn AI tự nhìn ảnh (slide, trang quét), người dùng cần chọn model nhận ảnh. Không được đoán nội dung của ảnh." };
        else toolImages.push(...imgs);
      }
      o.onEvent({ type: "tool_end", call, result, error });
      push(toolMsg(call, result));
      if (error) {
        const sig = `${call.name}|${error}`;
        let args = "";
        try {
          args = JSON.stringify(call.args ?? {});
        } catch {
          args = "?";
        }
        const before = prevFails.get(sig);
        const now = { run: (before?.run || 0) + 1, sameArgs: before && before.args === args ? before.sameArgs + 1 : 1, args };
        // Cùng lỗi lặp trong một bước chỉ tính một lần cho bước đó
        if (!stepFails.has(sig)) stepFails.set(sig, now);
        if (!stuck && (now.sameArgs >= STUCK_SAME_ARGS || now.run >= STUCK_SAME_ERROR)) stuck = { tool: call.name, error, times: now.run };
      } else if (specMap.get(call.name)?.write && !(result && typeof result === "object" && result.ok === false)) {
        writes++;
      }
    }
    prevFails = stepFails;
    if (toolImages.length) {
      push({
        role: "user",
        content: `[Add-in gửi kèm ${toolImages.length} ảnh do công cụ vừa chụp: ${toolImages.map((x) => x.label || "ảnh").join(", ")}. Đây là dữ liệu để bạn xem, không phải yêu cầu mới của người dùng.]`,
        images: toolImages.map((x) => ({ mime: x.mime, data: x.data })),
        fromTool: true,
      });
    }
    if (stuck) {
      // Dừng sau khi mọi lệnh của bước đã có kết quả (lịch sử vẫn hợp lệ)
      o.onEvent({ type: "stuck", tool: stuck.tool, error: stuck.error, times: stuck.times, step, writes });
      push({
        role: "assistant",
        content: `[Add-in dừng lượt: công cụ ${stuck.tool} báo cùng một lỗi ${stuck.times} bước liên tiếp (“${clip(stuck.error, 300)}”). Chờ người dùng quyết định làm tiếp theo cách khác hay thôi.]`,
      });
      return msgs;
    }
  }
  o.onEvent({ type: "limit", maxSteps, writes });
  push({ role: "assistant", content: `[Đã đạt giới hạn ${maxSteps} bước của một lượt; chờ người dùng bảo làm tiếp.]` });
  return msgs;
}
