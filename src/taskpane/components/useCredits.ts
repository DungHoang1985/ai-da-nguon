/** Hạn mức của key OpenRouter cho thẻ trong Cài đặt và dòng cảnh báo ở khung chat.
 *  Kết quả giữ trong bộ nhớ của khung (không ghi kho lưu), gọi lại thưa nhất 60 giây một lần trừ khi bấm "Làm mới". */
import * as React from "react";
import { Provider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { KeyInfo, fetchKeyInfo, hasKeyInfo } from "../../core/credits";
import { describeError } from "../../core/errors";

interface Cached {
  info?: KeyInfo;
  /** undefined = chưa thử /credits; false = key không được xem số dư */
  balanceAllowed?: boolean;
  error?: string;
  at: number;
}

const MIN_GAP_MS = 60000;
const cache: Record<string, Cached> = {};
const listeners = new Set<() => void>();
const notify = () => listeners.forEach((fn) => fn());

/** Đăng nhập, đăng xuất hoặc đổi key: bỏ kết quả cũ (key mới có thể có quyền khác) */
export function invalidateKeyInfo(providerId: string): void {
  delete cache[providerId];
  notify();
}

async function refresh(p: Provider, storage: Storage, force: boolean): Promise<void> {
  const key = storage.getKey(p.id);
  if (!key) {
    delete cache[p.id];
    notify();
    return;
  }
  const cur = cache[p.id];
  if (!force && cur && Date.now() - cur.at < MIN_GAP_MS) return;
  // Đánh dấu ngay để hai nơi cùng hỏi không gọi trùng
  cache[p.id] = { ...(cur || {}), at: Date.now() };
  try {
    const r = await fetchKeyInfo(p, key, { tryBalance: cur?.balanceAllowed !== false });
    cache[p.id] = { info: r.info, balanceAllowed: cur?.balanceAllowed === false ? false : r.balanceAllowed, at: Date.now() };
  } catch (e) {
    cache[p.id] = { ...(cur || {}), error: describeError(e, { provider: p }).text, at: Date.now() };
  }
  notify();
}

export function useCredits(storage: Storage, p: Provider | undefined, o: { auto?: boolean } = {}) {
  const [, bump] = React.useReducer((n: number) => n + 1, 0);
  const [loading, setLoading] = React.useState(false);
  const enabled = !!p && hasKeyInfo(p);
  const pid = p?.id || "";
  const hasKey = enabled && !!storage.getKey(pid);

  React.useEffect(() => {
    listeners.add(bump);
    return () => {
      listeners.delete(bump);
    };
  }, []);

  const reload = React.useCallback(
    async (force = true) => {
      if (!p || !enabled) return;
      setLoading(true);
      try {
        await refresh(p, storage, force);
      } finally {
        setLoading(false);
      }
    },
    [p, enabled, storage]
  );

  React.useEffect(() => {
    if (o.auto !== false && hasKey) reload(false);
  }, [pid, hasKey]); // eslint-disable-line react-hooks/exhaustive-deps

  const c = enabled ? cache[pid] : undefined;
  return { enabled, info: hasKey ? c?.info : undefined, balanceAllowed: c?.balanceAllowed, error: hasKey ? c?.error : undefined, loading, reload };
}
