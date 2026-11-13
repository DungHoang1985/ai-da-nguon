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
}

export interface StepResult {
  text: string;
  toolCalls: ToolCall[];
  raw?: { kind: string; parts?: any[] };
  finishReason?: string;
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
  systemPrompt(overview: any): string;
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
  | { type: "limit"; maxSteps: number };
