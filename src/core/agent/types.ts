/** Kiểu dữ liệu chung cho chế độ tác tử (AI tự gọi công cụ thao tác trên tài liệu). */

export interface ToolCall {
  id: string;
  name: string;
  args: any;
  /** id do chính nhà cung cấp trả về (Gemini cần gửi lại đúng id này) */
  nativeId?: boolean;
}

export type AgentMessage =
  | { role: "system"; content: string }
  | { role: "user"; content: string; images?: { mime: string; data: string }[] }
  | {
      role: "assistant";
      content: string;
      toolCalls?: ToolCall[];
      /** Nội dung gốc của nhà cung cấp, để gửi lại nguyên vẹn (vd. chữ ký suy nghĩ của Gemini) */
      raw?: { kind: string; parts?: any[] };
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
  /** Hạn chờ riêng cho công cụ này (mili giây); không ghi = hạn chờ chung của vòng lặp */
  timeoutMs?: number;
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
  raw?: { kind: string; parts?: any[] };
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
  execute(name: string, args: any, ctx: { log: UndoEntry[] }): Promise<any>;
  overview(): Promise<any>;
  /** Lời nhắc hệ thống CỐ ĐỊNH của ứng dụng (không chứa trạng thái tài liệu) — để bộ nhớ đệm lời nhắc có tác dụng */
  systemPrompt(): string;
  /** Trạng thái tài liệu lúc gửi, đưa vào tin nhắn của lượt (không vào lời nhắc hệ thống) */
  stateText?(overview: any): string;
  /** Ảnh chụp nội dung gọn để chia sẻ sang ứng dụng khác */
  snapshot?(): Promise<{ title?: string; data: any }>;
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
  | { type: "limit"; maxSteps: number }
  /** Nhà cung cấp quá tải hoặc hết hạn mức tạm thời: sắp thử lại lần `attempt` sau `waitMs` */
  | { type: "retry"; attempt: number; max: number; waitMs: number; reason: string }
  /** Số token của bước vừa xong */
  | { type: "usage"; step: number; usage: Usage }
  /** Câu trả lời bị cắt vì hết độ dài; `skipped` = số lệnh công cụ không được chạy; `gaveUp` = dừng vì bị cắt liên tiếp */
  | { type: "truncated"; skipped: number; gaveUp?: boolean };
