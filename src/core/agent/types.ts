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
  execute(name: string, args: any, ctx: { log: UndoEntry[]; said?: string; request?: string }): Promise<any>;
  overview(): Promise<any>;
  /** Lời nhắc hệ thống CỐ ĐỊNH của ứng dụng (không chứa trạng thái tài liệu) — để bộ nhớ đệm lời nhắc có tác dụng */
  systemPrompt(): string;
  /** Trạng thái tài liệu lúc gửi, đưa vào tin nhắn của lượt (không vào lời nhắc hệ thống) */
  stateText?(overview: any): string;
  /** Ảnh chụp nội dung gọn để chia sẻ sang ứng dụng khác */
  snapshot?(): Promise<{ title?: string; data: any }>;
  /** Chữ hiện tại của chỗ mà một lệnh ghi sắp sửa (để lượt soát lập luận đối chiếu); rỗng = không lấy được */
  reviewContext?(name: string, args: any): Promise<string>;
}

/** Công cụ dùng chung mọi ứng dụng (use_skill, read_other_app…) */
export interface ExtraTools {
  specs: ToolSpec[];
  execute(name: string, args: any): Promise<any>;
}

export type AgentEvent =
  | { type: "text"; delta: string }
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
  | { type: "limit"; maxSteps: number }
  /** Nhà cung cấp quá tải hoặc hết hạn mức tạm thời: sắp thử lại lần `attempt` sau `waitMs` */
  | { type: "retry"; attempt: number; max: number; waitMs: number; reason: string }
  /** Số token của bước vừa xong */
  | { type: "usage"; step: number; usage: Usage }
  /** Câu trả lời bị cắt vì hết độ dài; `skipped` = số lệnh công cụ không được chạy; `gaveUp` = dừng vì bị cắt liên tiếp */
  | { type: "truncated"; skipped: number; gaveUp?: boolean };
