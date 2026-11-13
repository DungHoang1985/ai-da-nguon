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
  | { role: "user"; content: string }
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
}

export type AgentEvent =
  | { type: "text"; delta: string }
  | { type: "step"; step: number }
  | { type: "tool_start"; call: ToolCall }
  | { type: "tool_end"; call: ToolCall; result: any; error?: string }
  | { type: "limit"; maxSteps: number };
