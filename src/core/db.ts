/** Kho khóa–giá trị bất đồng bộ: IndexedDB (dung lượng lớn), dự phòng bộ nhớ tạm. */

export interface AsyncKV {
  get<T = any>(key: string): Promise<T | undefined>;
  set(key: string, value: any): Promise<void>;
  del(key: string): Promise<void>;
  keys(prefix?: string): Promise<string[]>;
}

export class MemoryKV implements AsyncKV {
  private m = new Map<string, any>();
  async get<T>(key: string) {
    const v = this.m.get(key);
    return v === undefined ? undefined : (JSON.parse(JSON.stringify(v)) as T);
  }
  async set(key: string, value: any) {
    this.m.set(key, JSON.parse(JSON.stringify(value)));
  }
  async del(key: string) {
    this.m.delete(key);
  }
  async keys(prefix = "") {
    return [...this.m.keys()].filter((k) => k.startsWith(prefix));
  }
}

class IdbKV implements AsyncKV {
  private dbp: Promise<IDBDatabase>;
  constructor(name = "aidn", private store = "kv") {
    this.dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(name, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(store);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  private async tx<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest): Promise<T> {
    const db = await this.dbp;
    return new Promise<T>((resolve, reject) => {
      const t = db.transaction(this.store, mode);
      const r = fn(t.objectStore(this.store));
      r.onsuccess = () => resolve(r.result as T);
      r.onerror = () => reject(r.error);
    });
  }
  get<T>(key: string) {
    return this.tx<T | undefined>("readonly", (s) => s.get(key));
  }
  async set(key: string, value: any) {
    await this.tx("readwrite", (s) => s.put(value, key));
  }
  async del(key: string) {
    await this.tx("readwrite", (s) => s.delete(key));
  }
  async keys(prefix = "") {
    const all = await this.tx<IDBValidKey[]>("readonly", (s) => s.getAllKeys());
    return all.map(String).filter((k) => k.startsWith(prefix));
  }
}

let shared: AsyncKV | null = null;
export function getKV(): AsyncKV {
  if (shared) return shared;
  try {
    if (typeof indexedDB !== "undefined") {
      shared = new IdbKV();
      return shared;
    }
  } catch {
    /* dùng bộ nhớ tạm */
  }
  shared = new MemoryKV();
  return shared;
}
