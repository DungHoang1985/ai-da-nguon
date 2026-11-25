/** Kiểu dữ liệu chung cho chế độ tác tử (AI tự gọi công cụ thao tác trên tài liệu). */

export interface ToolCall {
  id: string;
  name: string;
  args: any;
  /** id do chính nhà cung cấp trả về (Gemini cần gửi lại đúng id này) */
  nativeId?: boolean;
}

/** Tệp gửi nguyên vẹn cho model tự đọc (PDF quét không có lớp chữ) */
export interface FilePart {
  mime: string;
  /** base64, không có tiền tố data: */
  data: string;
  name: string;
}

/** Nội dung gốc của nhà cung cấp, để gửi lại nguyên vẹn ở các bước sau (lịch sử chỉ thêm: không bao giờ sửa sau khi ghi).
 *  kind "gemini": parts (kèm chữ ký suy nghĩ) · kind "anthropic": blocks = thứ tự các khối của câu trả lời, trong đó khối
 *  thinking / redacted_thinking giữ nguyên văn (kèm signature), khối text giữ nguyên chữ của riêng nó, khối tool_use là mốc vị trí ·
 *  kind "openrouter": reasoning = reasoning_details của câu trả lời */
export interface RawAssistant {
  kind: string;
  parts?: any[];
  blocks?: any[];
  reasoning?: any[];
}

/** Tin nhắn trợ lý mang khối suy nghĩ có chữ ký (Anthropic `thinking`, OpenRouter `reasoning_details`): chỗ nào buộc phải đổi
 *  phần hội thoại đứng trước (tóm tắt, rút gọn khi lưu) thì phải bỏ `raw` này. Gemini không thuộc nhóm này. */
export function hasThinkingRaw(m: { role: string; raw?: RawAssistant }): boolean {
  return m.role === "assistant" && (m.raw?.kind === "anthropic" || m.raw?.kind === "openrouter");
}

export type AgentMessage =
  | { role: "system"; content: string }
  | { role: "user"; content: string; images?: { mime: string; data: string }[]; files?: FilePart[]; /** Tin do add-in thêm để gửi ảnh công cụ vừa chụp (không phải lời người dùng) */ fromTool?: boolean }
  | {
      role: "assistant";
      content: string;
      toolCalls?: ToolCall[];
      raw?: RawAssistant;
    }
  | { role: "tool"; toolCallId: string; name: string; content: string };

export interface ToolSpec {
  name: string;
  description: string;
  /** JSON Schema (tập con dùng được cho cả OpenAI, Anthropic, Gemini) */
  parameters: any;
  /** Công cụ có thay đổi tài liệu */
  write?: boolean;
  /** Trả về lời cảnh báo nếu lần gọi này rủi ro (ghi đè dữ liệu, xóa…) → luôn hỏi người dùng */
  risky?: (args: any) => Promise<string | null>;
  /** Cảnh báo của `risky` LUÔN được hỏi, kể cả khi người dùng tắt "hỏi trước thao tác rủi ro" (ghi sang ứng dụng khác) */
  alwaysAsk?: boolean;
  /** Hạn chờ riêng cho công cụ này (mili giây); không ghi = hạn chờ chung của vòng lặp */
  timeoutMs?: number;
  /** Công cụ ghi chữ vào điều khoản: khi bật "Soát lập luận trước khi sửa", lần gọi được một lượt model riêng soát trước khi chạy */
  reviewable?: boolean;
  /** Công cụ ghi làm đổi hình thức (bố cục slide): xong phải xem lại bằng công cụ `look` trước khi báo xong (cổng cuối lượt 4.12.0) */
  needsLook?: boolean;
  /** Công cụ xem lại hình thức (ảnh slide, kiểm bài): gọi sau các lệnh `needsLook` là đã kiểm */
  look?: boolean;
}

/** Số token của một lần gọi model (đã quy về cùng nghĩa cho mọi nhà cung cấp). */
export interface Usage {
  /** Tổng token đầu vào, gồm cả phần đọc từ và ghi vào bộ nhớ đệm */
  input?: number;
  output?: number;
  /** Token đầu vào đọc từ bộ nhớ đệm lời nhắc (rẻ hơn nhiều) */
  cacheRead?: number;
  /** Token đầu vào ghi vào bộ nhớ đệm */
  cacheWrite?: number;
}

export interface StepResult {
  text: string;
  toolCalls: ToolCall[];
  raw?: RawAssistant;
  finishReason?: string;
  /** Câu trả lời bị cắt vì hết độ dài đầu ra: lệnh gọi công cụ (nếu có) có thể dở dang, không được chạy */
  truncated?: boolean;
  usage?: Usage;
}

export interface UndoEntry {
  label: string;
  undo: () => Promise<void>;
}

/** Bộ công cụ của một ứng dụng Office */
export interface HostToolSet {
  host: string;
  specs: ToolSpec[];
  /** said = lời model nói ở bước đó, request = yêu cầu của người dùng ở lượt đó (cho nhật ký soát lại của Excel; công cụ không đưa hai thứ này trở lại model) */
  execute(name: string, args: any, ctx: { log: UndoEntry[]; said?: string; request?: string; verifier?: Verifier }): Promise<any>;
  overview(): Promise<any>;
  /** Lời nhắc hệ thống CỐ ĐỊNH của ứng dụng (không chứa trạng thái tài liệu) — để bộ nhớ đệm lời nhắc có tác dụng */
  systemPrompt(): string;
  /** Trạng thái tài liệu lúc gửi, đưa vào tin nhắn của lượt (không vào lời nhắc hệ thống) */
  stateText?(overview: any): string;
  /** Ảnh chụp nội dung gọn để chia sẻ sang ứng dụng khác */
  snapshot?(): Promise<{ title?: string; data: any }>;
  /** Chữ hiện tại của chỗ mà một lệnh ghi sắp sửa (để lượt soát lập luận đối chiếu); rỗng = không lấy được */
  reviewContext?(name: string, args: any): Promise<string>;
  /** Gọi khi một lượt tác tử kết thúc (xong, dừng hay lỗi) — Word chụp chữ các đoạn để lượt sau báo chỗ người dùng tự sửa (4.12.1) */
  afterTurn?(): Promise<void>;
}

/** Công cụ dùng chung mọi ứng dụng (use_skill, read_other_app…) */
/** Người soát ngữ cảnh mới cho công cụ soát lập luận (4.13.0): một lượt model riêng, không thấy hội thoại. `instructions`, `skills` là
 *  chỉ dẫn riêng và kỹ năng đã nạp của lượt — để người soát đối chiếu căn cứ loại user_instruction / skill_instruction. */
export interface Verifier {
  ask(system: string, user: string): Promise<string>;
  instructions?: string;
  skills?: { name: string; body: string }[];
}

export interface ExtraTools {
  specs: ToolSpec[];
  execute(name: string, args: any): Promise<any>;
}

export type AgentEvent =
  | { type: "text"; delta: string }
  /** Số đo token đầu vào của lần gửi đầu tiên của lượt (Đợt 12) — chỉ để hiện */
  | { type: "input_size"; tokens: number; context?: number }
  | { type: "step"; step: number }
  | { type: "tool_start"; call: ToolCall }
  | { type: "tool_end"; call: ToolCall; result: any; error?: string }
  /** Model đang suy nghĩ (chưa có chữ trả lời) */
  | { type: "thinking" }
  /** Đang soát lập luận của một lệnh sửa trước khi ghi */
  | { type: "review_start"; call: ToolCall }
  | { type: "review_end"; call: ToolCall; ok: boolean; note?: string }
  /** Nhà cung cấp từ chối tham số suy nghĩ: đã gửi lại không kèm phần đó */
  | { type: "thinking_off"; reason: string }
  /** Chạm trần số bước của lượt; `writes` = số lệnh ghi đã chạy thành công */
  | { type: "limit"; maxSteps: number; writes: number }
  /** Dừng sớm: một công cụ báo cùng một lỗi nhiều bước liên tiếp */
  | { type: "stuck"; tool: string; error: string; times: number; step: number; writes: number }
  /** Cổng cuối lượt: model định kết thúc khi còn lệnh lỗi chưa xử lý, add-in nhắc và cho trả lời tiếp một lần */
  | { type: "error_gate"; tools: string[] }
  /** Cổng cuối lượt thứ hai: đổi hình thức (bố cục slide) mà chưa xem lại, add-in nhắc và cho làm tiếp một lần */
  | { type: "look_gate"; tools: string[] }
  /** Nhà cung cấp quá tải hoặc hết hạn mức tạm thời: sắp thử lại lần `attempt` sau `waitMs` */
  | { type: "retry"; attempt: number; max: number; waitMs: number; reason: string }
  /** Số token của bước vừa xong */
  | { type: "usage"; step: number; usage: Usage }
  /** Câu trả lời bị cắt vì hết độ dài; `skipped` = số lệnh công cụ không được chạy; `gaveUp` = dừng vì bị cắt liên tiếp */
  | { type: "truncated"; skipped: number; gaveUp?: boolean }
  /** Chế độ Miễn phí: chuyển sang nguồn kế tiếp (Đợt 9). `partial` = chữ dở của nguồn trước đã hiện, không vào lịch sử */
  | { type: "switch"; from: string; to: string; reason: string; partial: boolean };
