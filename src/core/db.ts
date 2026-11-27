/** Kho khóa–giá trị bất đồng bộ: IndexedDB (dung lượng lớn), dự phòng bộ nhớ tạm. */

export interface AsyncKV {
  get<T = any>(key: string): Promise<T | undefined>;
  set(key: string, value: any): Promise<void>;
  del(key: string): Promise<void>;
  keys(prefix?: string): Promise<string[]>;
  /** Đọc – sửa – ghi một khóa trong MỘT giao dịch (B11): hai khung cùng sửa không ghi đè mất phần của nhau.
   *  `fn` phải đồng bộ (IndexedDB tự đóng giao dịch khi chờ việc ngoài). Trả giá trị đã ghi. */
  update<T = any>(key: string, fn: (cur: T | undefined) => T): Promise<T>;
}

const clone = <T>(v: T): T => (v === undefined ? v : JSON.parse(JSON.stringify(v)));

export class MemoryKV implements AsyncKV {
  private m = new Map<string, any>();
  async get<T>(key: string) {
    return clone(this.m.get(key)) as T | undefined;
  }
  async set(key: string, value: any) {
    this.m.set(key, clone(value));
  }
  async del(key: string) {
    this.m.delete(key);
  }
  async keys(prefix = "") {
    return [...this.m.keys()].filter((k) => k.startsWith(prefix));
  }
  async update<T>(key: string, fn: (cur: T | undefined) => T) {
    const next = fn(clone(this.m.get(key)));
    this.m.set(key, clone(next));
    return clone(next);
  }
}

/** Mở IndexedDB; mọi kết quả chỉ báo xong khi giao dịch đã ghi xong (`oncomplete`), không ở `onsuccess` của lệnh (B11). */
export class IdbKV implements AsyncKV {
  readonly ready: Promise<IDBDatabase>;
  constructor(
    private idb: IDBFactory,
    name = "aidn",
    private store = "kv"
  ) {
    this.ready = new Promise((resolve, reject) => {
      const req = idb.open(name, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(store);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
      // onblocked: tạm thời (khung khác đang giữ bản cũ), mở thường xong sau đó — không coi là hỏng
    });
  }
  private async run<T>(mode: IDBTransactionMode, body: (s: IDBObjectStore, done: (v: T) => void, fail: (e: unknown) => void) => void): Promise<T> {
    const db = await this.ready;
    return new Promise<T>((resolve, reject) => {
      const t = db.transaction(this.store, mode);
      let out: T;
      let firstErr: unknown;
      const fail = (e: unknown) => {
        if (firstErr === undefined) firstErr = e ?? new Error("Lỗi IndexedDB");
        try {
          t.abort();
        } catch {
          /* đã đóng */
        }
        reject(firstErr);
      };
      t.oncomplete = () => resolve(out);
      t.onerror = () => reject(firstErr ?? t.error);
      t.onabort = () => reject(firstErr ?? t.error ?? new Error("Giao dịch IndexedDB bị hủy"));
      try {
        body(t.objectStore(this.store), (v) => (out = v), fail);
      } catch (e) {
        fail(e);
      }
    });
  }
  get<T>(key: string) {
    return this.run<T | undefined>("readonly", (s, done) => {
      const r = s.get(key);
      r.onsuccess = () => done(r.result as T);
    });
  }
  async set(key: string, value: any) {
    await this.run<void>("readwrite", (s) => void s.put(value, key));
  }
  async del(key: string) {
    await this.run<void>("readwrite", (s) => void s.delete(key));
  }
  async keys(prefix = "") {
    const all = await this.run<IDBValidKey[]>("readonly", (s, done) => {
      const r = s.getAllKeys();
      r.onsuccess = () => done(r.result);
    });
    return all.map(String).filter((k) => k.startsWith(prefix));
  }
  update<T>(key: string, fn: (cur: T | undefined) => T) {
    return this.run<T>("readwrite", (s, done, fail) => {
      const r = s.get(key);
      r.onsuccess = () => {
        let next: T;
        try {
          next = fn(r.result as T | undefined);
        } catch (e) {
          // Lỗi trong fn: hủy giao dịch, trả đúng lỗi gốc (onabort chỉ có lỗi rỗng)
          fail(e);
          return;
        }
        s.put(next, key);
        done(next);
      };
    });
  }
}

/** Dùng IndexedDB; mở lỗi thì chuyển hẳn sang bộ nhớ tạm (không hỏng mãi) và báo qua `onFallback` (B11). */
export class FallbackKV implements AsyncKV {
  private cur: AsyncKV;
  private fellBack = false;
  private listeners: ((why: string) => void)[] = [];
  constructor(make: () => IdbKV) {
    try {
      const idb = make();
      this.cur = idb;
      idb.ready.catch((e) => this.fall(e));
    } catch (e) {
      this.cur = new MemoryKV();
      this.fall(e);
    }
  }
  get degraded() {
    return this.fellBack;
  }
  onFallback(cb: (why: string) => void): () => void {
    if (this.fellBack) cb(this.why);
    this.listeners.push(cb);
    return () => (this.listeners = this.listeners.filter((x) => x !== cb));
  }
  private why = "";
  private fall(e: unknown) {
    if (this.fellBack) return;
    this.fellBack = true;
    this.why = e instanceof Error ? e.message : String(e ?? "");
    if (!(this.cur instanceof MemoryKV)) this.cur = new MemoryKV();
    this.listeners.forEach((cb) => cb(this.why));
  }
  private async call<R>(f: (kv: AsyncKV) => Promise<R>): Promise<R> {
    const kv = this.cur;
    try {
      return await f(kv);
    } catch (e) {
      // Chỉ chuyển khi chính việc MỞ cơ sở dữ liệu hỏng; lỗi của một lệnh thì trả cho người gọi
      // (lệnh có thể đã bắt đầu trước khi biết mở hỏng: vẫn chuyển rồi chạy lại trên bộ nhớ tạm)
      if (kv instanceof IdbKV && (await kv.ready.then(() => false, () => true))) {
        this.fall(e);
        return f(this.cur);
      }
      throw e;
    }
  }
  get<T>(key: string) {
    return this.call((kv) => kv.get<T>(key));
  }
  set(key: string, value: any) {
    return this.call((kv) => kv.set(key, value));
  }
  del(key: string) {
    return this.call((kv) => kv.del(key));
  }
  keys(prefix?: string) {
    return this.call((kv) => kv.keys(prefix));
  }
  update<T>(key: string, fn: (cur: T | undefined) => T) {
    return this.call((kv) => kv.update<T>(key, fn));
  }
}

let shared: FallbackKV | null = null;
export function getKV(): FallbackKV {
  if (shared) return shared;
  shared = new FallbackKV(() => {
    if (typeof indexedDB === "undefined") throw new Error("Không có IndexedDB");
    return new IdbKV(indexedDB);
  });
  return shared;
}