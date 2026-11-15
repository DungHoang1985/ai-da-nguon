/** Lưu cục bộ trên máy (localStorage của add-in). Không gửi đi đâu ngoài nhà cung cấp tương ứng. */
import type { ModelInfo } from "./models";
import { cleanKey } from "./errors";

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
  /** Bật tìm kiếm web (theo nhà cung cấp) */
  webSearch: boolean;
  /** Hướng dẫn riêng theo từng ứng dụng */
  instructions: Record<string, string>;
}

export const DEFAULT_SETTINGS: AppSettings = {
  agentMode: true,
  askBeforeWrite: false,
  toolsOnly: true,
  confirmRisky: true,
  shareContext: true,
  webSearch: false,
  instructions: {},
};

export interface Selection {
  providerId?: string;
  /** Bản cũ (≤3.5): một model chung cho mọi nguồn */
  modelId?: string;
  freeOnly?: boolean;
  /** Từ 3.6: model đã dùng lần trước theo từng nguồn */
  models?: Record<string, string>;
}

export class Storage {
  constructor(private store: KVStore = defaultStore()) {}

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

  getKeys(): Record<string, string> {
    return this.read<Record<string, string>>("keys", {});
  }

  getKey(providerId: string): string {
    // Làm sạch cả key đã lưu từ bản cũ (có thể lẫn ký tự ẩn khi dán)
    return cleanKey(this.getKeys()[providerId] || "");
  }

  setKey(providerId: string, key: string): void {
    const keys = this.getKeys();
    keys[providerId] = cleanKey(key);
    this.write("keys", keys);
  }

  removeKey(providerId: string): void {
    const keys = this.getKeys();
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
