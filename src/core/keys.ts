/** Phím trong ô nhập của khung chat (Đợt 0A bước 5) — hàm thuần để kiểm thử.
 *  Enter gửi · Shift+Enter xuống dòng · Ctrl+Enter vẫn gửi · bộ gõ đang ghép chữ (IME) thì không làm gì
 *  · danh sách "/" đang mở: mũi tên di chuyển, Enter/Tab chọn, Esc đóng · đang chạy: Enter không gửi, Esc = Dừng. */

export type InputKeyAction = "send" | "newline" | "blocked" | "menu-next" | "menu-prev" | "menu-pick" | "menu-close" | "stop" | "none";

export interface KeyInfo {
  key: string;
  shiftKey?: boolean;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
  /** KeyboardEvent.isComposing: bộ gõ (vd. Telex của Windows) đang ghép chữ */
  isComposing?: boolean;
  /** 229 = trình duyệt báo phím đang được bộ gõ xử lý */
  keyCode?: number;
}

export function inputKeyAction(k: KeyInfo, st: { menuOpen: boolean; busy: boolean }): InputKeyAction {
  if (k.isComposing || k.keyCode === 229) return "none";
  if (st.menuOpen) {
    if (k.key === "ArrowDown") return "menu-next";
    if (k.key === "ArrowUp") return "menu-prev";
    if ((k.key === "Enter" || k.key === "Tab") && !k.shiftKey) return "menu-pick";
    if (k.key === "Escape") return "menu-close";
  }
  if (k.key === "Escape") return st.busy ? "stop" : "none";
  if (k.key === "Enter") {
    if (k.shiftKey || k.altKey) return "newline";
    return st.busy ? "blocked" : "send";
  }
  return "none";
}
