/** Chép vào bộ nhớ tạm và tải tệp ra đĩa từ khung add-in, có lối dự phòng cho WebView2 trong Office. */

/** Chép chữ: thử Clipboard API trước, không được thì dùng execCommand("copy") (WebView2 trong Office có thể chặn Clipboard API). */
export async function copyText(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    /* thử cách dự phòng */
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch {
    return false;
  }
}

/** Tải tệp ra đĩa (thư mục Tải xuống). Trả về false nếu môi trường không cho tải. */
export function downloadText(fileName: string, text: string, mime = "application/json"): boolean {
  try {
    const blob = new Blob([text], { type: `${mime};charset=utf-8` });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    return true;
  } catch {
    return false;
  }
}
