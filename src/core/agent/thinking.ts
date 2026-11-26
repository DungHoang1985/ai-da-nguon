/** Công tắc "Suy nghĩ kỹ": tham số suy nghĩ theo từng nhà cung cấp và từng đời model.
 *  Căn cứ: tài liệu chính thức đọc ngày 15/11/2026, tóm tắt ở tasks/dot/dot-1b/bang-chung/05-tra-tai-lieu-nha-cung-cap.md.
 *  Các tham số này KHÔNG nằm trong lời nhắc hệ thống, bộ công cụ hay tin nhắn, nên bật / tắt không làm vỡ bộ nhớ đệm lời nhắc
 *  và không làm hỏng khối suy nghĩ đã có trong lịch sử. */
import { Provider, quirksOf } from "../providers";

/** Cách một model Claude nhận tham số suy nghĩ:
 *  - bound: suy nghĩ bật sẵn, chỉ kiểu "adaptive"; khối suy nghĩ gắn với phần hội thoại đứng trước (Fable/Mythos 5.x, Opus 5.5,
 *    Sonnet 5.5, Haiku 5.5 và model ra sau)
 *  - adaptive_on: suy nghĩ bật sẵn, kiểu "adaptive" (Opus 5, Sonnet 5)
 *  - adaptive_off: không gửi thì không suy nghĩ; bật bằng "adaptive" (Opus 4.7, 4.8)
 *  - adaptive_46: như trên, không có mức xhigh (Opus 4.6, Sonnet 4.6)
 *  - budget: "enabled" kèm budget_tokens (Haiku 4.5, Sonnet 4.5, Opus 4.5 trở về 3.7)
 *  - none: không suy nghĩ (Claude 3.5 trở về trước) */
export type AnthropicThinking = "bound" | "adaptive_on" | "adaptive_off" | "adaptive_46" | "budget" | "none";

export function anthropicThinkingKind(model: string): AnthropicThinking {
  const m = String(model || "").toLowerCase();
  if (/claude-(2|instant)|claude-3-(haiku|opus|sonnet)|claude-3-5/.test(m)) return "none";
  if (/claude-3-7|claude-(opus|sonnet)-4-(0|1|5)(?!\d)|claude-(opus|sonnet)-4-20\d{6}|claude-haiku-4-5/.test(m)) return "budget";
  if (/claude-(opus|sonnet)-4-6/.test(m)) return "adaptive_46";
  if (/claude-opus-4-[78]/.test(m)) return "adaptive_off";
  if (/claude-(opus|sonnet)-5(?!-5)/.test(m)) return "adaptive_on";
  return "bound";
}

/** Tiêu đề beta cho phép yêu cầu máy chủ tự bỏ khối suy nghĩ không còn hợp lệ thay vì báo lỗi 400 */
export const THINKING_BINDING_BETA = "thinking-binding-controls-2026-08-01";
export const THINKING_BUDGET_TOKENS = 16000;

export interface AnthropicThinkingParams {
  thinking?: any;
  output_config?: any;
  beta?: string;
}

/** Tham số suy nghĩ gửi cho Anthropic. on = công tắc "Suy nghĩ kỹ". */
export function anthropicThinkingParams(model: string, on: boolean, maxTokens: number): AnthropicThinkingParams {
  const kind = anthropicThinkingKind(model);
  if (kind === "bound") {
    // Luôn kèm drop_block: lịch sử của add-in là chỉ thêm, nhưng lời nhắc hệ thống hoặc bộ công cụ có thể đổi giữa hội thoại
    // (cập nhật add-in, nhập thêm kỹ năng, bật tìm web). Khi đó máy chủ tự bỏ khối suy nghĩ cũ thay vì từ chối cả yêu cầu.
    const out: AnthropicThinkingParams = {
      thinking: { type: "adaptive", block_binding: { prefix_mismatch_behavior: "drop_block" } },
      beta: THINKING_BINDING_BETA,
    };
    // Mức mặc định là medium ở Opus 5.5 và Haiku 5.5, high ở các model còn lại: bật công tắc thì lên một mức
    if (on) out.output_config = { effort: /claude-(opus|haiku)-5-5/.test(model.toLowerCase()) ? "high" : "xhigh" };
    return out;
  }
  if (!on) return {};
  if (kind === "adaptive_on") return { thinking: { type: "adaptive" }, output_config: { effort: "xhigh" } };
  if (kind === "adaptive_off" || kind === "adaptive_46") return { thinking: { type: "adaptive" } };
  if (kind === "budget") {
    const budget = Math.min(THINKING_BUDGET_TOKENS, Math.floor(maxTokens) - 4096);
    return budget >= 1024 ? { thinking: { type: "enabled", budget_tokens: budget } } : {};
  }
  return {};
}

/** Tham số suy nghĩ cho một nguồn kiểu Anthropic, có tính phương ngữ của nguồn (Đợt 6):
 *  - nguồn gốc (không có dialect): như anthropicThinkingParams;
 *  - "none": không gửi gì;
 *  - "enabled": chỉ bật / tắt (DeepSeek: adaptive coi như enabled, budget_tokens bị bỏ qua — tài liệu 17/11/2026);
 *  - "anthropic" + bare: theo đời model Claude nhưng không kèm tiêu đề beta, nên bỏ luôn block_binding (tham số cần tiêu đề đó). */
export function dialectThinkingParams(provider: Provider, model: string, on: boolean, maxTokens: number): AnthropicThinkingParams {
  const d = provider.dialect;
  // Mã model của nơi trung chuyển có tiền tố nhà cung cấp (anthropic/claude-…): bỏ tiền tố trước khi xét đời model
  const bareModel = String(model || "").replace(/^[^/]+\//, "").replace(/\./g, "-");
  if (!d) return anthropicThinkingParams(model, on, maxTokens);
  const mode = d.thinking || "anthropic";
  if (mode === "none") return {};
  if (mode === "enabled") return { thinking: { type: on ? "enabled" : "disabled" } };
  const p = anthropicThinkingParams(bareModel, on, maxTokens);
  if (!d.bare) return p;
  const out: AnthropicThinkingParams = {};
  if (p.thinking) out.thinking = p.thinking.block_binding ? { type: p.thinking.type } : p.thinking;
  if (p.output_config) out.output_config = p.output_config;
  return out;
}

/** generationConfig.thinkingConfig của Gemini khi bật công tắc; model không rõ thì không gửi gì */
export function geminiThinkingConfig(model: string): any | undefined {
  const m = String(model || "").toLowerCase();
  if (/gemini-2\.5-pro/.test(m)) return { thinkingBudget: 32768 };
  if (/gemini-2\.5/.test(m)) return { thinkingBudget: 24576 };
  if (/gemini-([3-9]|\d{2})(\.|-|$)/.test(m)) return { thinkingLevel: "high" };
  return undefined;
}

/** Công tắc "Suy nghĩ kỹ" có tác dụng với nguồn và model này không (để khung chat ghi rõ) */
export function thinkingToggleWorks(provider: Provider, model: string): boolean {
  if (provider.kind === "anthropic") {
    const mode = provider.dialect?.thinking || "anthropic";
    if (mode === "none") return false;
    if (mode === "enabled") return true;
    return anthropicThinkingKind(String(model || "").replace(/^[^/]+\//, "").replace(/\./g, "-")) !== "none";
  }
  if (provider.kind === "gemini") return !!geminiThinkingConfig(model);
  return !!quirksOf(provider).reasoning;
}

/** Lỗi 400 do nhà cung cấp không nhận tham số suy nghĩ hoặc khối suy nghĩ gửi lại: gửi lại một lần không kèm các phần đó */
export function isThinkingParamError(e: any): boolean {
  const status = e?.status;
  if (status !== 400 && status !== 422) return false;
  return /thinking|signature|block_binding|reasoning|effort|output_config|thought|anthropic-beta|beta header|unexpected value\(s\).*beta/i.test(String(e?.message || ""));
}

/** Nguồn nhận nguyên tệp PDF để model tự đọc (PDF quét không có lớp chữ) */
export function pdfCapable(provider: Provider): boolean {
  if (provider.kind === "anthropic") return !provider.dialect || provider.dialect.pdf === true;
  return provider.kind === "gemini" || !!quirksOf(provider).pdf;
}
