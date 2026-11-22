/** Lịch sử hội thoại lưu trên máy (IndexedDB), theo từng ứng dụng. */
import { AsyncKV } from "./db";
import { AgentMessage } from "./agent/types";

export interface ConversationMeta {
  id: string;
  host: string;
  title: string;
  updatedAt: number;
}

export interface Conversation extends ConversationMeta {
  /** Nội dung hiển thị (không gồm hàm Hoàn tác) */
  items: any[];
  /** Hội thoại gửi model */
  history: AgentMessage[];
  /** Đợt 12 phần E: khối bộ nhớ chụp lúc hội thoại bắt đầu (giữ nguyên lời nhắc hệ thống khi mở lại) */
  memory?: import("./memory").MemoryItem[];
}

export const MAX_CONVERSATIONS = 60;
const STORED_TOOL_CHARS = 4000;

export function newConversationId(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

export function titleFrom(text: string): string {
  const t = String(text || "")
    .replace(/\s+/g, " ")
    .trim();
  return t.length > 60 ? t.slice(0, 57) + "…" : t || "Hội thoại mới";
}

/** Bỏ phần nặng trước khi lưu: kết quả công cụ dài, ảnh và tệp đính kèm.
 *  Khối suy nghĩ của Anthropic gắn với phần hội thoại đứng trước nó: từ tin nhắn đầu tiên bị rút gọn trở đi, bản lưu không còn
 *  khớp với hội thoại đã gửi nên các khối suy nghĩ phía sau cũng bỏ (mở lại hội thoại thì không gửi lại khối đã hết hợp lệ). */
export function slimForStorage(history: AgentMessage[]): AgentMessage[] {
  let changed = false;
  return history.map((m) => {
    if (m.role === "tool" && m.content.length > STORED_TOOL_CHARS) {
      changed = true;
      return { ...m, content: m.content.slice(0, STORED_TOOL_CHARS) + "… [đã rút gọn khi lưu]" };
    }
    if (m.role === "user" && (m.images?.length || m.files?.length)) {
      changed = true;
      const { images, files, ...rest } = m;
      const notes = [
        images?.length ? `[Đã đính kèm ${images.length} ảnh]` : "",
        files?.length ? `[Đã đính kèm tệp ${files.map((f) => `“${f.name}”`).join(", ")} (gửi nguyên tệp); bản lưu không giữ tệp, cần hỏi tiếp về tệp thì đính kèm lại]` : "",
      ].filter(Boolean);
      return { ...rest, content: `${rest.content}\n${notes.join("\n")}` };
    }
    if (changed && m.role === "assistant" && m.raw?.kind === "anthropic") {
      const { raw: _raw, ...rest } = m;
      return rest;
    }
    return m;
  });
}

/** Khóa của "chế độ thử" trong sessionStorage của khung: bài thử tự động bật để hội thoại thử không vào lịch sử
 *  (lịch sử có trần, hội thoại thử từng đẩy mất hội thoại thật). Không có nút trên giao diện; mất khi đóng Office. */
export const TEST_MODE_KEY = "aidn-che-do-thu";

export function testModeOn(store?: Pick<Storage, "getItem">): boolean {
  try {
    const s = store ?? (typeof sessionStorage !== "undefined" ? sessionStorage : undefined);
    return s?.getItem(TEST_MODE_KEY) === "1";
  } catch {
    return false;
  }
}

export class ConversationStore {
  /** noSave() trả true thì save() không ghi gì (chế độ thử) */
  constructor(
    private kv: AsyncKV,
    private noSave?: () => boolean
  ) {}

  /** Hội thoại sẽ bị xóa khỏi lịch sử nếu lưu thêm MỘT hội thoại mới (lịch sử đang đầy); không có thì undefined */
  async nextToDrop(currentId?: string): Promise<ConversationMeta | undefined> {
    if (this.noSave?.()) return undefined;
    const index = ((await this.kv.get<ConversationMeta[]>("conv-index")) || []).slice().sort((a, b) => b.updatedAt - a.updatedAt);
    if (currentId && index.some((x) => x.id === currentId)) return undefined;
    return index.length >= MAX_CONVERSATIONS ? index[index.length - 1] : undefined;
  }

  async list(host?: string): Promise<ConversationMeta[]> {
    const index = (await this.kv.get<ConversationMeta[]>("conv-index")) || [];
    return index.filter((c) => !host || c.host === host).sort((a, b) => b.updatedAt - a.updatedAt);
  }

  async get(id: string): Promise<Conversation | undefined> {
    return this.kv.get<Conversation>("conv:" + id);
  }

  async save(c: Conversation): Promise<void> {
    if (this.noSave?.()) return;
    const stored: Conversation = { ...c, history: slimForStorage(c.history), updatedAt: Date.now() };
    await this.kv.set("conv:" + c.id, stored);
    let index = (await this.kv.get<ConversationMeta[]>("conv-index")) || [];
    index = index.filter((x) => x.id !== c.id);
    index.unshift({ id: c.id, host: c.host, title: c.title, updatedAt: stored.updatedAt });
    index.sort((a, b) => b.updatedAt - a.updatedAt);
    const drop = index.slice(MAX_CONVERSATIONS);
    index = index.slice(0, MAX_CONVERSATIONS);
    for (const d of drop) await this.kv.del("conv:" + d.id);
    await this.kv.set("conv-index", index);
  }

  async delete(id: string): Promise<void> {
    await this.kv.del("conv:" + id);
    const index = ((await this.kv.get<ConversationMeta[]>("conv-index")) || []).filter((x) => x.id !== id);
    await this.kv.set("conv-index", index);
  }

  async clear(host?: string): Promise<void> {
    const index = (await this.kv.get<ConversationMeta[]>("conv-index")) || [];
    const keep: ConversationMeta[] = [];
    for (const c of index) {
      if (!host || c.host === host) await this.kv.del("conv:" + c.id);
      else keep.push(c);
    }
    await this.kv.set("conv-index", keep);
  }
}
