/**
 * Khối xử lý sự kiện của một lượt (trước đây là `onEvent` ~90 dòng trong Chat.tsx — phần để sau của nhóm 2, làm 27/11/2026).
 * Thuần: không gọi React, không gọi Office. Nhận sự kiện của runTurn, giữ bộ đếm của lượt (lệnh ghi thành công / thất bại,
 * công cụ đã gọi, vị trí cổng cuối lượt, đang thử lại) và trả danh sách việc cho khung chat làm theo đúng thứ tự:
 * bỏ thông báo thử lại → sửa mục trả lời → nối chữ → đặt thông báo.
 */
import { addUsage } from "./agent/turn";
import { Usage } from "./agent/types";
import { writeSucceeded } from "./claims";
import { SHOW_SIZE_FROM, sizeNote } from "./inputsize";

/** Các trường của mục trả lời mà khối sự kiện đọc / sửa (kiểu đầy đủ ở Chat.tsx) */
export interface TurnItem {
  text: string;
  segments?: any[];
  wrote?: boolean;
  usage?: Usage;
  calls?: number;
  answeredBy?: string;
  inputEst?: number;
  firstInput?: number;
  paused?: { steps: number; writes: number; stuck?: string };
}

type Item = TurnItem;

/** Các khối chữ của câu trả lời (để cắt phần trước lúc cổng cuối lượt nhắc AI — 4.11.15) */
export const textParts = (it: Item): string[] =>
  it.segments ? it.segments.filter((sg) => sg.type === "text").map((sg: any) => sg.text) : it.text ? [it.text] : [];

export type TurnEffect<T extends TurnItem = TurnItem> =
  | { kind: "clear_retry" }
  | { kind: "patch"; fn: (it: T) => T }
  | { kind: "text"; text: string }
  | { kind: "notice"; text: string };

export interface TurnEventDeps {
  /** Lượt có xoay vòng nguồn (Miễn phí): báo cỡ lượt không kèm cửa sổ ngữ cảnh của một model */
  rotation: () => boolean;
  /** Tên công cụ là lệnh ghi */
  isWrite: (name: string) => boolean;
  /** Nhãn của một lệnh công cụ trên thẻ bước */
  describe: (call: { name: string; args?: any }) => string;
}

export class TurnEventTracker<T extends TurnItem = TurnItem> {
  /** Công cụ đã gọi trong lượt — để nhận ra ca "nói đã sửa mà không ghi" */
  readonly called: string[] = [];
  /** Số lệnh ghi không thành công */
  failedWrites = 0;
  /** Tham số của các lệnh ghi thành công — đối chiếu với các chỗ "A" → "B" AI kể đã sửa (vòng thử 3, 24/11/2026) */
  readonly writtenArgs: string[] = [];
  /** Độ dài chữ lúc cổng cuối lượt nhắc AI; undefined = cổng không chạy. Chữ sau cổng nối tiếp vào khối cuối nên cắt theo ký tự. */
  gateCut: number | undefined;
  /** Đang hiện thông báo "đang thử lại" */
  retrying = false;
  private needSep = false;

  constructor(private deps: TurnEventDeps) {}

  /** Việc bỏ thông báo thử lại (chỉ khi đang hiện) */
  private clear(out: TurnEffect<T>[]) {
    if (!this.retrying) return;
    this.retrying = false;
    out.push({ kind: "clear_retry" });
  }

  handle(ev: any): TurnEffect<T>[] {
    const out: TurnEffect<T>[] = [];
    const patch = (fn: (it: T) => T) => out.push({ kind: "patch", fn });
    switch (ev?.type) {
      case "input_size":
        // Đợt 12: lượt lớn thì báo số đo (chỉ để hiện, không vào lịch sử)
        patch((it) => ({ ...it, inputEst: ev.tokens }));
        if (ev.tokens >= SHOW_SIZE_FROM) out.push({ kind: "notice", text: sizeNote(ev.tokens, this.deps.rotation() ? undefined : ev.context) });
        break;
      case "compacted":
        out.push({ kind: "notice", text: "Cuộc trò chuyện dài: đã tự tóm tắt phần đầu để AI nhớ gọn hơn." });
        break;
      case "step":
        if (ev.step > 1) this.needSep = true;
        patch((it) => ({ ...it, calls: Math.max(it.calls || 0, ev.step) }));
        break;
      case "text":
        this.clear(out);
        out.push({ kind: "text", text: this.needSep ? "\n" + ev.delta : ev.delta });
        this.needSep = false;
        break;
      case "tool_start":
        this.clear(out);
        patch((it) => ({
          ...it,
          segments: [...(it.segments || []), { type: "step", id: ev.call.id, label: this.deps.describe(ev.call), status: "running" } as any],
        }));
        break;
      case "tool_end": {
        const isWrite = this.deps.isWrite(ev.call.name);
        this.called.push(ev.call.name);
        // Công cụ ghi báo lỗi, trả ok=false hoặc "thay 0 chỗ" thì tài liệu chưa đổi
        const done = isWrite && writeSucceeded(ev.result, ev.error);
        if (isWrite && !done) this.failedWrites++;
        if (done) {
          try {
            this.writtenArgs.push(JSON.stringify(ev.call.args ?? {}));
          } catch {
            /* tham số không chuyển được thành chữ: bỏ qua */
          }
        }
        patch((it) => ({
          ...it,
          wrote: it.wrote || done,
          segments: (it.segments || []).map((sg: any) =>
            sg.type === "step" && sg.id === ev.call.id
              ? {
                  ...sg,
                  status: ev.error ? "error" : "ok",
                  detail: ev.error || ev.result?.message,
                  brief: !ev.error && typeof ev.result?.brief === "string" ? ev.result.brief : undefined,
                }
              : sg
          ),
        }));
        break;
      }
      case "review_start":
        this.clear(out);
        patch((it) => ({
          ...it,
          segments: [
            ...(it.segments || []),
            { type: "step", id: "review:" + ev.call.id, label: `Soát lập luận trước khi ghi: ${this.deps.describe(ev.call)}`, status: "running" } as any,
          ],
        }));
        break;
      case "review_end":
        patch((it) => ({
          ...it,
          segments: (it.segments || []).map((sg: any) =>
            sg.type === "step" && sg.id === "review:" + ev.call.id
              ? { ...sg, status: ev.ok ? "ok" : "error", detail: ev.ok ? ev.note || "Đạt" : `Chưa ghi — ${ev.note || "có vấn đề"}`, brief: ev.ok ? ev.note || "đạt" : undefined }
              : sg
          ),
        }));
        break;
      case "switch":
        this.clear(out);
        patch((it) => ({ ...it, answeredBy: ev.to }));
        out.push({ kind: "text", text: `\n↪ ${ev.from}: ${ev.reason} — đổi sang ${ev.to}${ev.partial ? " (trả lời lại bước này)" : ""}.\n` });
        this.needSep = false;
        break;
      case "thinking_off":
        out.push({ kind: "notice", text: "Nguồn đang dùng không nhận chế độ suy nghĩ với model này: đã gửi lại yêu cầu ở chế độ thường." });
        break;
      case "limit":
        patch((it) => ({ ...it, paused: { steps: ev.maxSteps, writes: ev.writes } }));
        break;
      case "error_gate":
        // Ghi vị trí lúc cổng nhắc: đọc trên mục hiện tại khi áp việc (giữ nguyên mục)
        patch((it) => {
          this.gateCut = textParts(it).join("\n").length;
          return it;
        });
        break;
      case "stuck":
        patch((it) => ({ ...it, paused: { steps: ev.step, writes: ev.writes, stuck: `${ev.tool} báo cùng một lỗi ${ev.times} bước liên tiếp` } }));
        break;
      case "retry": {
        this.retrying = true;
        const why = String(ev.reason || "");
        out.push({
          kind: "notice",
          text: `${why.charAt(0).toUpperCase() + why.slice(1)}: đang thử lại lần ${ev.attempt}/${ev.max} sau ${Math.round(ev.waitMs / 1000)} giây… Bấm Dừng để hủy.`,
        });
        break;
      }
      case "usage":
        this.clear(out);
        patch((it) => ({
          ...it,
          usage: addUsage(it.usage || {}, ev.usage),
          ...(it.firstInput === undefined && ev.step === 1 && ev.usage?.input ? { firstInput: ev.usage.input } : {}),
        }));
        break;
      case "truncated":
        if (!ev.skipped) out.push({ kind: "text", text: "\n(Câu trả lời bị cắt vì quá dài. Gửi “tiếp tục” để AI viết tiếp.)" });
        else if (ev.gaveUp)
          out.push({
            kind: "text",
            text: "\n(Dừng lượt: câu trả lời bị cắt vì quá dài nhiều lần liên tiếp, các lệnh ghi dở chưa được thực hiện. Hãy nhờ AI làm từng phần nhỏ hơn.)",
          });
        break;
    }
    return out;
  }
}
