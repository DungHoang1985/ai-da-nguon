/** Chia sẻ ảnh chụp nội dung giữa Excel – Word – PowerPoint (cùng bộ nhớ cục bộ của add-in). */
import { KVStore } from "./storage";

export const SHARE_PREFIX = "aidn.share.";
export const SHARE_MAX_AGE_MS = 30 * 60 * 1000;
const MAX_SHARE_CHARS = 200000;

export interface Snapshot {
  host: string;
  at: number;
  title?: string;
  data: any;
}

function store(s?: KVStore): KVStore | null {
  try {
    return s || ((globalThis as any).localStorage as KVStore) || null;
  } catch {
    return null;
  }
}

export function publishSnapshot(host: string, data: any, title?: string, s?: KVStore): void {
  const st = store(s);
  if (!st) return;
  let json = JSON.stringify({ host, at: Date.now(), title, data } as Snapshot);
  if (json.length > MAX_SHARE_CHARS) {
    json = JSON.stringify({ host, at: Date.now(), title, data: { truncated: true, text: JSON.stringify(data).slice(0, MAX_SHARE_CHARS - 500) } });
  }
  try {
    st.setItem(SHARE_PREFIX + host, json);
  } catch {
    /* đầy bộ nhớ → bỏ qua */
  }
}

export function readSnapshot(host: string, now = Date.now(), s?: KVStore): Snapshot | null {
  const st = store(s);
  if (!st) return null;
  try {
    const raw = st.getItem(SHARE_PREFIX + host);
    if (!raw) return null;
    const snap = JSON.parse(raw) as Snapshot;
    return now - snap.at > SHARE_MAX_AGE_MS ? null : snap;
  } catch {
    return null;
  }
}

/** Thực thi công cụ read_other_app */
export function readOtherApp(app: string, current: string, s?: KVStore): any {
  const map: Record<string, string> = { excel: "Excel", word: "Word", powerpoint: "PowerPoint", ppt: "PowerPoint" };
  const host = map[String(app || "").toLowerCase()] || app;
  if (host === current) throw new Error(`Đang ở ${current}; dùng công cụ của chính ứng dụng này để đọc.`);
  const snap = readSnapshot(host, Date.now(), s);
  if (!snap) {
    throw new Error(
      `Chưa có dữ liệu từ ${host}. Hãy nhờ người dùng mở ${host} cùng tệp cần dùng và bật khung AI Đa Nguồn ở đó (dữ liệu tự chia sẻ khi khung đang mở).`
    );
  }
  return {
    ok: true,
    app: host,
    title: snap.title,
    updated: new Date(snap.at).toLocaleString("vi-VN"),
    data: snap.data,
  };
}
