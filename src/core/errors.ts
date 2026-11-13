/** Đổi lỗi kỹ thuật thành thông báo tiếng Việt dễ hiểu. */
export class ApiError extends Error {
  constructor(
    message: string,
    public status?: number
  ) {
    super(message);
  }
}

export function messageForStatus(status: number, detail = ""): string {
  const d = detail ? ` (${detail.slice(0, 200)})` : "";
  if (status === 401 || status === 403) return `Key không đúng hoặc không có quyền${d}`;
  if (status === 402) return `Tài khoản hết tiền hoặc cần nạp thêm${d}`;
  if (status === 404) return `Không tìm thấy model hoặc địa chỉ API${d}`;
  if (status === 429) return `Hết hạn mức (gọi quá nhiều). Thử lại sau hoặc đổi model${d}`;
  if (status >= 500) return `Máy chủ nhà cung cấp đang lỗi (${status}). Thử lại sau${d}`;
  return `Lỗi ${status}${d}`;
}

export function networkErrorMessage(e: unknown): string {
  if (e instanceof ApiError) return e.message;
  if (e instanceof DOMException && e.name === "AbortError") return "Đã dừng.";
  if (e instanceof TypeError) {
    return "Không kết nối được. Có thể mất mạng, hoặc nhà cung cấp này chặn gọi trực tiếp từ add-in (CORS).";
  }
  return (e as Error)?.message || String(e);
}

export async function ensureOk(res: Response): Promise<void> {
  if (res.ok) return;
  let detail = "";
  try {
    const text = await res.text();
    try {
      const j = JSON.parse(text);
      detail = j?.error?.message || j?.message || j?.error || text;
      if (typeof detail !== "string") detail = JSON.stringify(detail);
    } catch {
      detail = text;
    }
  } catch {
    /* bỏ qua */
  }
  throw new ApiError(messageForStatus(res.status, detail), res.status);
}
