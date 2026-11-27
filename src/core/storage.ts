/** Lưu cục bộ trên máy (localStorage của add-in). Không gửi đi đâu ngoài nhà cung cấp tương ứng. */
import type { ModelInfo } from "./models";
import { cleanKey } from "./errors";
import { allProviders, CustomSource, setCustomSources as registerCustomSources } from "./providers";
import { MemoryItem, MemoryScope, cleanMemoryText, newMemoryId, sanitizeMemory, validateMemoryText } from "./memory";

export interface KVStore {
  getItem(k: string): string | null;
  setItem(k: string, v: string): void;
  removeItem(k: string): void;
}

const memory: Record<string, string> = {};
const memoryStore: KVStore = {
  getItem: (k) => (k in memory ? memory[k] : null),
  setItem: (k, v) => {
    memory[k] = v;
  },
  removeItem: (k) => {
    delete memory[k];
  },
};

function defaultStore(): KVStore {
  try {
    const ls = (globalThis as any).localStorage as KVStore | undefined;
    if (ls) {
      ls.setItem("__aidn_test", "1");
      ls.removeItem("__aidn_test");
      return ls;
    }
  } catch {
    /* localStorage không dùng được → dùng bộ nhớ tạm */
  }
  return memoryStore;
}

const PREFIX = "aidn.";

export interface AppSettings {
  /** Cho AI tự thao tác trên tài liệu (gọi công cụ) */
  agentMode: boolean;
  /** Hỏi trước mỗi thao tác sửa tài liệu */
  askBeforeWrite: boolean;
  /** OpenRouter: chỉ hiện model gọi được công cụ */
  toolsOnly: boolean;
  /** Luôn hỏi trước thao tác rủi ro: ghi đè ô có dữ liệu, xóa vùng/đoạn/slide */
  confirmRisky: boolean;
  /** Chia sẻ nội dung tài liệu đang mở cho các ứng dụng Office khác */
  shareContext: boolean;
  /** Khung này nhận lệnh ghi do khung ở ứng dụng Office khác gửi sang (người dùng luôn được hỏi ở khung gửi) */
  crossWrite: boolean;
  /** Kết nối với từng ứng dụng khác (đọc chéo và ghi chéo); thiếu khóa = bật */
  crossApps: Record<string, boolean>;
  /** Bật tìm kiếm web (theo nhà cung cấp) */
  webSearch: boolean;
  /** Suy nghĩ kỹ: cho model suy nghĩ nhiều hơn trước khi trả lời (chậm và tốn hơn; tùy nhà cung cấp và model) */
  deepThink: boolean;
  /** Word: soát lập luận của mỗi lệnh ghi chữ vào điều khoản bằng một lượt model riêng trước khi ghi */
  reviewEdits: boolean;
  /** Excel: ghi mỗi thao tác ghi của AI vào sheet "Nhật ký AI" (mặc định tắt; sheet đi theo tệp khi gửi cho người khác) */
  excelLog: boolean;
  /** Excel: tô nền các ô AI vừa ghi (mặc định tắt) */
  excelHighlight: boolean;
  /** Hướng dẫn riêng theo từng ứng dụng */
  instructions: Record<string, string>;
}

export const DEFAULT_SETTINGS: AppSettings = {
  agentMode: true,
  askBeforeWrite: false,
  toolsOnly: true,
  confirmRisky: true,
  shareContext: true,
  crossWrite: true,
  crossApps: {},
  webSearch: false,
  deepThink: false,
  reviewEdits: false,
  excelLog: false,
  excelHighlight: false,
  instructions: {},
};

export interface Selection {
  providerId?: string;
  /** Bản cũ (≤3.5): một model chung cho mọi nguồn */
  modelId?: string;
  freeOnly?: boolean;
  /** Từ 3.6: model đã dùng lần trước theo từng nguồn */
  models?: Record<string, string>;
  /** Đợt 9: chế độ Miễn phí / Trả phí; chưa có = chưa chọn → khung hiện thẻ chọn chế độ */
  mode?: "free" | "paid";
  /** Đợt 9: thứ tự xoay vòng nguồn miễn phí (mã nguồn) */
  freeOrder?: string[];
  /** Đợt 9: nguồn đã chọn lần cuối ở từng chế độ — đổi qua lại chế độ thì về đúng nguồn cũ (soát mã Đợt 9, M4) */
  providerByMode?: { free?: string; paid?: string };
}

export class Storage {
  constructor(private store: KVStore = defaultStore()) {
    this.reloadCustomSources();
  }

  /** B22: đăng ký nguồn tự khai vào danh sách nguồn — tường minh lúc khởi động và sau khi khôi phục sao lưu
   *  (trước đây chỉ chạy kèm getKeys(), ai gọi allProviders() trước thì không thấy nguồn tự khai) */
  reloadCustomSources(): void {
    registerCustomSources(this.getCustomSources());
  }

  private read<T>(key: string, fallback: T): T {
    try {
      const raw = this.store.getItem(PREFIX + key);
      return raw ? (JSON.parse(raw) as T) : fallback;
    } catch {
      return fallback;
    }
  }

  private write(key: string, value: unknown): void {
    try {
      this.store.setItem(PREFIX + key, JSON.stringify(value));
    } catch {
      /* bỏ qua */
    }
  }

  /** Key đúng như đã lưu (không kèm key dùng chung) */
  private rawKeys(): Record<string, string> {
    return this.read<Record<string, string>>("keys", {});
  }

  /** Key theo nguồn. Nguồn dùng chung key với nguồn khác (`keyFrom`) nhận key của nguồn đó — chỉ tính ra, không lưu thêm bản sao. */
  getKeys(): Record<string, string> {
    const keys = { ...this.rawKeys() };
    for (const p of allProviders()) {
      if (!p.keyFrom) continue;
      if (keys[p.keyFrom]) keys[p.id] = keys[p.keyFrom];
      else delete keys[p.id];
    }
    return keys;
  }

  /** Các nguồn kiểu Anthropic do người dùng tự khai */
  getCustomSources(): CustomSource[] {
    const list = this.read<CustomSource[]>("customSources", []);
    return Array.isArray(list) ? list : [];
  }

  setCustomSources(list: CustomSource[]): void {
    this.write("customSources", list);
    registerCustomSources(list);
  }

  getKey(providerId: string): string {
    // Làm sạch cả key đã lưu từ bản cũ (có thể lẫn ký tự ẩn khi dán)
    return cleanKey(this.getKeys()[providerId] || "");
  }

  setKey(providerId: string, key: string): void {
    const keys = this.rawKeys();
    keys[providerId] = cleanKey(key);
    this.write("keys", keys);
  }

  removeKey(providerId: string): void {
    const keys = this.rawKeys();
    delete keys[providerId];
    this.write("keys", keys);
    const cache = this.getModelCache();
    delete cache[providerId];
    this.write("models", cache);
  }

  getModelCache(): Record<string, { at: number; models: ModelInfo[] }> {
    return this.read("models", {});
  }

  setModels(providerId: string, models: ModelInfo[], at = Date.now()): void {
    const cache = this.getModelCache();
    cache[providerId] = { at, models };
    this.write("models", cache);
  }

  getSettings(): AppSettings {
    return { ...DEFAULT_SETTINGS, ...this.read<Partial<AppSettings>>("settings", {}) };
  }

  setSettings(s: Partial<AppSettings>): void {
    this.write("settings", { ...this.getSettings(), ...s });
  }

  /** Đợt 12 phần E: bộ nhớ dài hạn giữa các hội thoại */
  getMemory(): MemoryItem[] {
    return sanitizeMemory(this.read<unknown>("memory", []));
  }

  setMemory(list: MemoryItem[]): void {
    this.write("memory", sanitizeMemory(list));
  }

  /** Thêm một mục; trả câu lỗi nếu không thêm được */
  addMemory(text: string, scope: MemoryScope, source: "user" | "ai", now = Date.now()): { ok: true; item: MemoryItem } | { ok: false; error: string } {
    const t = cleanMemoryText(text);
    const list = this.getMemory();
    const err = validateMemoryText(t, list, scope);
    if (err) return { ok: false, error: err };
    const item: MemoryItem = { id: newMemoryId(now), text: t, scope, source, createdAt: now };
    this.setMemory([...list, item]);
    return { ok: true, item };
  }

  /** Đợt 9: thời điểm nghỉ của từng nguồn miễn phí (hết lượt hôm nay / tạm thời) */
  getCooldowns(): Record<string, { until: number; reason: string }> {
    const v = this.read<Record<string, { until: number; reason: string }>>("freeCooldowns", {});
    return v && typeof v === "object" ? v : {};
  }

  setCooldowns(v: Record<string, { until: number; reason: string }>): void {
    this.write("freeCooldowns", v);
  }

  getSelection(): Selection {
    return this.read<Selection>("selection", {});
  }

  setSelection(sel: Partial<Selection>): void {
    this.write("selection", { ...this.getSelection(), ...sel });
  }

  /** Model đã dùng lần trước ở nguồn này (bản cũ chỉ nhớ một model chung → coi là model của nguồn đang chọn khi đó). */
  getModelFor(providerId: string): string {
    const sel = this.getSelection();
    const byProvider = sel.models || {};
    if (byProvider[providerId]) return byProvider[providerId];
    if (!sel.models && sel.providerId === providerId && sel.modelId) return sel.modelId;
    return "";
  }

  setModelFor(providerId: string, modelId: string): void {
    const sel = this.getSelection();
    const models = { ...(sel.models || {}) };
    if (!sel.models && sel.providerId && sel.modelId) models[sel.providerId] = sel.modelId;
    if (modelId) models[providerId] = modelId;
    else delete models[providerId];
    this.write("selection", { ...sel, models });
  }
}
