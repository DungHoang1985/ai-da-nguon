/** Hạn mức của key OpenRouter cho thẻ trong Cài đặt và dòng cảnh báo ở khung chat.
 *  Kết quả giữ trong bộ nhớ của khung (không ghi kho lưu, không giữ key), gọi lại thưa nhất 60 giây một lần trừ khi bấm "Làm mới". */
import * as React from "react";
import { Provider } from "../../core/providers";
import { Storage } from "../../core/storage";
import { KeyInfo, fetchKeyInfo, hasKeyInfo } from "../../core/credits";
import { describeError } from "../../core/errors";

interface Cached {
  info?: KeyInfo;
  /** undefined = chưa biết; false = máy chủ nói rõ key này không được xem số dư */
  balanceAllowed?: boolean;
  error?: string;
  at: number;
}

const MIN_GAP_MS = 60000;
const cache: Record<string, Cached> = {};
/** Tăng mỗi lần key của một nguồn đổi: kết quả của lượt gọi bắt đầu trước đó (hạn mức của key cũ) bị bỏ */
const generation: Record<string, number> = {};
const listeners = new Set<() => void>();
const notify = () => listeners.forEach((fn) => fn());

/** Đăng nhập, đăng xuất hoặc đổi key: bỏ kết quả cũ; các thẻ đang hiện tự đọc lại cho key mới */
export function invalidateKeyInfo(providerId: string): void {
  generation[providerId] = (generation[providerId] || 0) + 1;
  delete cache[providerId];
  notify();
}

async function refresh(p: Provider, storage: Storage, force: boolean): Promise<void> {
  const key = storage.getKey(p.id);
  if (!key) {
    if (cache[p.id]) {
      delete cache[p.id];
      notify();
    }
    return;
  }
  const cur = cache[p.id];
  if (!force && cur && Date.now() - cur.at < MIN_GAP_MS) return;
  const gen = generation[p.id] || 0;
  // Đánh dấu ngay để hai nơi cùng hỏi không gọi trùng
  cache[p.id] = { ...(cur || {}), at: Date.now() };
  let next: Cached;
  try {
    const r = await fetchKeyInfo(p, key, { tryBalance: cur?.balanceAllowed !== false });
    next = { info: r.info, balanceAllowed: r.balanceAllowed ?? cur?.balanceAllowed, at: Date.now() };
  } catch (e) {
    next = { ...(cur || {}), error: describeError(e, { provider: p }).text, at: Date.now() };
  }
  // Key đã đổi trong lúc chờ: kết quả này là của key cũ
  if ((generation[p.id] || 0) !== gen) return;
  cache[p.id] = next;
  notify();
}

export function useCredits(storage: Storage, p: Provider | undefined, o: { auto?: boolean } = {}) {
  const [tick, bump] = React.useReducer((n: number) => n + 1, 0);
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

  // Lần đầu, và mỗi khi kết quả bị bỏ vì key đổi (đăng nhập lại): tự đọc lại. Có kết quả (kể cả lỗi) thì không gọi nữa.
  React.useEffect(() => {
    if (o.auto !== false && hasKey && !cache[pid]) reload(false);
  }, [pid, hasKey, tick]); // eslint-disable-line react-hooks/exhaustive-deps

  const c = enabled ? cache[pid] : undefined;
  return { enabled, info: hasKey ? c?.info : undefined, balanceAllowed: c?.balanceAllowed, error: hasKey ? c?.error : undefined, loading, reload };
}
