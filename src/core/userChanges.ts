/** Đoạn người dùng tự sửa giữa hai lượt (4.12.1, phỏng theo khối <user_changes> của bản chính thức — phân tích mã 25/11/2026).
 *  Cuối mỗi lượt add-in chụp chữ các đoạn theo định danh đoạn; đầu lượt sau so lại để báo AI chỗ người dùng đã tự sửa, thêm, xóa
 *  (AI tôn trọng, không ghi đè). Hàm thuần, kiểm thử được; phần Office ở wordTools.ts. Bản chụp chỉ nằm trong bộ nhớ của khung. */

export interface SnapRow {
  /** Định danh đoạn (uniqueLocalId, WordApi 1.6); thiếu thì bỏ qua đoạn đó */
  id?: string;
  /** Số thứ tự đoạn hiện tại */
  i: number;
  /** Chữ hiện tại của đoạn */
  text: string;
  /** Đoạn đã bị xóa có theo dõi (còn trong văn bản chờ duyệt) */
  deleted?: boolean;
}

export interface UserChanges {
  note: string;
  changed: { index: number; before: string; now: string }[];
  inserted: { index: number; text: string }[];
  removed: number;
  /** Có nhiều chỗ hơn số được liệt kê */
  more?: number;
}

export const MAX_LISTED = 8;
export const CLIP = 160;

const clip = (s: string) => (s.length > CLIP ? s.slice(0, CLIP) + "…" : s);

/** Bản chụp cuối lượt: định danh → chữ */
export function snapshotOf(rows: SnapRow[]): Map<string, string> {
  const m = new Map<string, string>();
  for (const r of rows) if (r.id && !r.deleted) m.set(r.id, r.text);
  return m;
}

/** So bản chụp cuối lượt trước với văn bản hiện tại. Không có gì đổi thì trả null. */
export function diffUserChanges(prev: Map<string, string> | null | undefined, rows: SnapRow[]): UserChanges | null {
  if (!prev || !prev.size) return null;
  const live = rows.filter((r) => r.id && !r.deleted);
  if (!live.length) return null;
  const seen = new Set<string>();
  const changed: UserChanges["changed"] = [];
  const inserted: UserChanges["inserted"] = [];
  let all = 0;
  for (const r of live) {
    const id = r.id as string;
    seen.add(id);
    if (!prev.has(id)) {
      // Đoạn trống mới (Enter thừa) không đáng báo
      if (!r.text.trim()) continue;
      all++;
      if (changed.length + inserted.length < MAX_LISTED) inserted.push({ index: r.i, text: clip(r.text) });
    } else if (prev.get(id) !== r.text) {
      all++;
      if (changed.length + inserted.length < MAX_LISTED) changed.push({ index: r.i, before: clip(prev.get(id) as string), now: clip(r.text) });
    }
  }
  let removed = 0;
  prev.forEach((text, id) => {
    if (!seen.has(id) && text.trim()) removed++;
  });
  all += removed;
  if (!all) return null;
  const listed = changed.length + inserted.length + (removed ? 1 : 0);
  return {
    note: "Người dùng TỰ sửa văn bản kể từ khi bạn xong lượt trước (đoạn đổi chữ, đoạn thêm, số đoạn đã xóa). Tôn trọng các thay đổi này: không ghi đè, không đưa về chữ cũ; nếu liên quan tới yêu cầu thì dựa vào chữ mới. Khi nhắc tới, nói rõ là người dùng đã sửa (“anh/chị đã sửa…”), không viết “tôi đã sửa”. Chỉ số đoạn là chỉ số hiện tại.",
    changed,
    inserted,
    removed,
    ...(all > listed ? { more: all - listed } : {}),
  };
}
