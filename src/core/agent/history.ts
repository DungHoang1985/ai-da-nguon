/** Chuẩn hóa lịch sử hội thoại trước khi gửi cho nguồn đích (khi đổi nhà cung cấp giữa hội thoại, mở lại hội thoại cũ, lượt bị dừng).
 *  Chỉ đổi trong bản gửi đi; lịch sử gốc giữ nguyên để quay lại nguồn cũ vẫn dùng được mã gốc (Gemini cần). */
import { Provider } from "../providers";
import { AgentMessage, ToolCall } from "./types";
import { pdfCapable } from "./thinking";

export interface HistoryTarget {
  provider: Provider;
  /** Model đích nhận ảnh: true = có, false = không, undefined = không rõ */
  vision?: boolean;
}

const ALNUM = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

/** Mã lời gọi công cụ hợp lệ với nguồn đích chưa.
 *  Mistral: đúng 9 ký tự chữ-số · Anthropic: chữ, số, _ - · OpenAI và nguồn kiểu OpenAI: tối đa 40 ký tự, chữ, số, _ -
 *  (OpenRouter cho thêm . và : để giữ dạng "functions.ten:0" của Kimi) · Gemini: chỉ gửi lại mã do chính Gemini cấp. */
export function callIdOk(p: Provider, id: string): boolean {
  if (p.id === "mistral") return /^[a-zA-Z0-9]{9}$/.test(id);
  if (p.kind === "gemini") return true;
  if (p.kind === "anthropic") return /^[a-zA-Z0-9_-]{1,64}$/.test(id);
  if (p.id === "openrouter") return /^[a-zA-Z0-9_.:-]{1,40}$/.test(id);
  return /^[a-zA-Z0-9_-]{1,40}$/.test(id);
}

/** Mã mới 9 ký tự chữ-số (hợp lệ với mọi nguồn), cố định theo mã gốc để mỗi lượt gửi ra cùng một mã. */
export function remapCallId(id: string, used: Set<string>): string {
  for (let salt = 0; ; salt++) {
    let h1 = 0x811c9dc5 ^ salt;
    let h2 = 0x01000193 + salt;
    const s = `${id}#${salt}`;
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      h1 = Math.imul(h1 ^ c, 0x01000193) >>> 0;
      h2 = Math.imul(h2 + c, 0x5bd1e995) >>> 0;
    }
    let out = "";
    let a = h1;
    let b = h2;
    for (let i = 0; i < 9; i++) {
      const src = i % 2 ? a : b;
      out += ALNUM[src % 62];
      if (i % 2) a = Math.floor(a / 62) ^ Math.imul(b, 31);
      else b = Math.floor(b / 62) ^ Math.imul(a, 17);
      a >>>= 0;
      b >>>= 0;
    }
    if (!used.has(out)) {
      used.add(out);
      return out;
    }
  }
}

const str = (v: unknown): string => (typeof v === "string" ? v : v === undefined || v === null ? "" : String(v));

/** Trả về bản lịch sử hợp lệ cho nguồn đích. Không sửa mảng/đối tượng gốc. */
export function prepareHistory(messages: AgentMessage[], target: HistoryTarget): AgentMessage[] {
  const p = target.provider;
  const turns = messages.filter((m) => m.role !== "system");

  // 1) Dữ liệu cũ/hỏng: nội dung undefined, tham số undefined
  const clean: AgentMessage[] = turns.map((m) => {
    if (m.role === "user") return { ...m, content: str(m.content) };
    if (m.role === "tool") return { ...m, content: str(m.content), toolCallId: str(m.toolCallId), name: str(m.name) };
    if (m.role === "assistant") {
      const content = str(m.content);
      return {
        ...m,
        // Chữ chỉ có khoảng trắng (vd. "\n\n" của vLLM) → coi như rỗng
        content: content.trim() ? content : "",
        toolCalls: (m.toolCalls || []).map((tc) => ({ ...tc, id: str(tc.id), name: str(tc.name), args: tc.args && typeof tc.args === "object" ? tc.args : {} })),
      };
    }
    return m;
  });

  // 2) Giữ đúng cặp gọi – kết quả: kết quả mồ côi bỏ đi; lời gọi thiếu kết quả (lượt bị dừng) thêm kết quả thay thế
  const paired: AgentMessage[] = [];
  for (let i = 0; i < clean.length; i++) {
    const m = clean[i];
    if (m.role === "tool") continue; // xử lý cùng lời gọi ở dưới
    paired.push(m);
    if (m.role === "assistant" && m.toolCalls?.length) {
      // gom các kết quả liền sau
      const results: AgentMessage[] = [];
      let j = i + 1;
      while (j < clean.length && clean[j].role === "tool") results.push(clean[j++]);
      for (const tc of m.toolCalls) {
        const r = results.find((x) => x.role === "tool" && x.toolCallId === tc.id);
        paired.push(
          r || { role: "tool", toolCallId: tc.id, name: tc.name, content: '{"ok":false,"error":"(không có kết quả: lượt trước bị dừng giữa chừng)"}' }
        );
      }
      i = j - 1;
    }
  }

  // 3) Mã lời gọi công cụ theo luật của nguồn đích (Gemini: giữ mã gốc do Gemini cấp)
  const used = new Set<string>();
  const map = new Map<string, string>();
  for (const m of paired) if (m.role === "assistant") for (const tc of m.toolCalls || []) if (callIdOk(p, tc.id)) used.add(tc.id);
  const idFor = (id: string): string => {
    if (callIdOk(p, id) && id) return id;
    if (!map.has(id)) map.set(id, remapCallId(id, used));
    return map.get(id)!;
  };
  let out: AgentMessage[] = paired.map((m) => {
    if (m.role === "assistant" && m.toolCalls?.length) {
      const toolCalls: ToolCall[] = m.toolCalls.map((tc) => ({ ...tc, id: idFor(tc.id) }));
      return { ...m, toolCalls };
    }
    if (m.role === "tool") return { ...m, toolCallId: idFor(m.toolCallId) };
    return m;
  });

  // 4) Ảnh: không nhận ảnh → bỏ hết; chưa rõ → chỉ giữ ảnh của câu hỏi cuối cùng
  const lastUser = out.map((m) => m.role).lastIndexOf("user");
  out = out.map((m, i) => {
    if (m.role !== "user" || !m.images?.length) return m;
    // Anthropic: mọi model Claude hiện hành đều nhận ảnh, và bỏ ảnh của lượt cũ là sửa tin nhắn cũ (làm hỏng khối suy nghĩ
    // đã gắn với phần hội thoại đứng trước, vỡ bộ nhớ đệm) → chưa rõ thì vẫn giữ
    const keep = target.vision === true || (target.vision === undefined && (i === lastUser || p.kind === "anthropic"));
    if (keep) return m;
    const { images, ...rest } = m;
    const note =
      target.vision === false
        ? `[${images.length} ảnh đính kèm không gửi được: model này không nhận ảnh]`
        : `[${images.length} ảnh đã gửi ở lượt trước, không gửi lại]`;
    return { ...rest, content: rest.content ? `${rest.content}\n${note}` : note };
  });

  // 4b) Tệp gửi nguyên vẹn (PDF quét): nguồn không nhận tệp → thay bằng một dòng ghi rõ, để model báo lại cho người dùng.
  //     Nguồn nhận tệp thì giữ ở mọi lượt (không bỏ ở lượt sau: bỏ là sửa tin nhắn cũ, làm vỡ bộ nhớ đệm lời nhắc).
  // Model đã biết là không nhận ảnh (chỉ đọc chữ) thì cũng không đọc được PDF quét: gửi tệp sẽ bị từ chối ở mọi lượt sau
  if (!pdfCapable(p) || target.vision === false) {
    const why = pdfCapable(p) ? `model đang chọn của nguồn ${p.name} chỉ đọc chữ, không đọc được PDF quét` : `nguồn ${p.name} không đọc được PDF quét`;
    out = out.map((m) => {
      if (m.role !== "user" || !m.files?.length) return m;
      const { files, ...rest } = m;
      const note = `[Tệp ${files.map((f) => `“${f.name}”`).join(", ")} là PDF quét (không có lớp chữ) nên KHÔNG gửi được: ${why}. Hãy nói rõ điều này với người dùng và gợi ý chuyển sang nguồn Anthropic hoặc Gemini (hoặc một model đọc được tệp) để đọc tệp.]`;
      return { ...rest, content: rest.content ? `${rest.content}\n${note}` : note };
    });
  }

  // 5) Mistral không nhận câu hỏi người dùng ngay sau kết quả công cụ → chèn một câu trả lời nối
  if (p.id === "mistral") {
    const fixed: AgentMessage[] = [];
    out.forEach((m, i) => {
      fixed.push(m);
      if (m.role === "tool" && out[i + 1]?.role === "user") fixed.push({ role: "assistant", content: "(đã nhận kết quả công cụ)" });
    });
    out = fixed;
  }
  return out;
}
