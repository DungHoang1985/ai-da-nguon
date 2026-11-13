/** Lưu cục bộ trên máy (localStorage của add-in). Không gửi đi đâu ngoài nhà cung cấp tương ứng. */
import type { ModelInfo } from "./models";

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
    return this.getKeys()[providerId] || "";
  }

  setKey(providerId: string, key: string): void {
    const keys = this.getKeys();
    keys[providerId] = key.trim();
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

  getSelection(): { providerId?: string; modelId?: string; freeOnly?: boolean } {
    return this.read("selection", {});
  }

  setSelection(sel: { providerId?: string; modelId?: string; freeOnly?: boolean }): void {
    this.write("selection", { ...this.getSelection(), ...sel });
  }
}
