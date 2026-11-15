/* global Office */
/** Hộp thoại của Office và mở trình duyệt ngoài (Đợt 2: đăng nhập).
 *  Đối chiếu learn.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins (bản 28/09/2026):
 *  trang đầu của hộp thoại phải cùng tên miền với khung và dùng HTTPS; sau đó được chuyển sang trang HTTPS khác;
 *  mỗi khung chỉ mở được một hộp thoại; không dùng window.open. */

export interface DialogHandle {
  close(): void;
}

export interface DialogHandlers {
  /** Chuỗi trang trong hộp thoại gửi về bằng messageParent */
  onMessage(message: string): void;
  /** Hộp thoại đóng không do add-in: người dùng đóng, hoặc trang không tải được */
  onClosed(why: "user" | "error", text: string): void;
}

export class DialogError extends Error {
  constructor(
    message: string,
    public code?: number
  ) {
    super(message);
    this.name = "DialogError";
  }
}

export function dialogSupported(): boolean {
  try {
    return typeof Office !== "undefined" && typeof Office.context?.ui?.displayDialogAsync === "function";
  } catch {
    return false;
  }
}

/** Câu tiếng Việt cho mã lỗi của displayDialogAsync (12004–12011) */
export function dialogOpenErrorText(code?: number, message = ""): string {
  if (code === 12007) return "Đang có một cửa sổ đăng nhập mở sẵn. Đóng cửa sổ đó rồi bấm lại.";
  if (code === 12009) return "Office đã chặn cửa sổ đăng nhập (anh chọn không cho mở). Bấm lại và chọn cho phép, hoặc dùng lối dán mã bên dưới.";
  if (code === 12011) return "Trình duyệt chặn cửa sổ bật lên. Dùng lối dán mã bên dưới.";
  if (code === 12004 || code === 12005) return `Office không cho mở trang đăng nhập (lỗi ${code}). Dùng lối dán mã bên dưới.`;
  return `Không mở được cửa sổ đăng nhập của Office${code ? ` (lỗi ${code})` : ""}${message ? `: ${message.slice(0, 120)}` : ""}. Dùng lối dán mã bên dưới.`;
}

/** Câu tiếng Việt cho sự kiện của hộp thoại đang mở (12002, 12003, 12006) */
export function dialogEventText(code?: number): { why: "user" | "error"; text: string } {
  if (code === 12006) return { why: "user", text: "Đã đóng cửa sổ đăng nhập trước khi xong." };
  if (code === 12002) return { why: "error", text: "Cửa sổ đăng nhập không tải được trang (mất mạng hoặc trang bị chặn). Dùng lối dán mã bên dưới." };
  if (code === 12003) return { why: "error", text: "Cửa sổ đăng nhập bị chuyển tới trang không an toàn nên Office đã dừng. Dùng lối dán mã bên dưới." };
  return { why: "error", text: `Cửa sổ đăng nhập gặp lỗi${code ? ` ${code}` : ""}. Dùng lối dán mã bên dưới.` };
}

/** Mở hộp thoại của Office tại `url` (trang của chính add-in). */
export function openDialog(url: string, handlers: DialogHandlers, size: { height?: number; width?: number } = {}): Promise<DialogHandle> {
  return new Promise((resolve, reject) => {
    if (!dialogSupported()) {
      reject(new DialogError("Bản Office này không có cửa sổ đăng nhập cho add-in. Dùng lối dán mã bên dưới."));
      return;
    }
    try {
      Office.context.ui.displayDialogAsync(url, { height: size.height ?? 75, width: size.width ?? 40, promptBeforeOpen: false }, (r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new DialogError(dialogOpenErrorText(r.error?.code, r.error?.message), r.error?.code));
          return;
        }
        const dialog = r.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
          if (typeof arg?.message === "string") handlers.onMessage(arg.message);
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
          const ev = dialogEventText(arg?.error);
          handlers.onClosed(ev.why, ev.text);
        });
        resolve({
          close: () => {
            try {
              dialog.close();
            } catch {
              /* đã đóng */
            }
          },
        });
      });
    } catch (e: any) {
      reject(new DialogError(dialogOpenErrorText(undefined, String(e?.message || e))));
    }
  });
}

/** Mở một địa chỉ trong trình duyệt mặc định của máy (ngoài Office). */
export function openInBrowser(url: string): boolean {
  try {
    const ui: any = typeof Office !== "undefined" ? Office.context?.ui : undefined;
    if (ui && typeof ui.openBrowserWindow === "function") {
      ui.openBrowserWindow(url);
      return true;
    }
  } catch {
    /* thử cách dưới */
  }
  try {
    // Có "noopener" thì window.open luôn trả null dù đã mở: không dựa vào giá trị trả về
    window.open(url, "_blank", "noopener,noreferrer");
    return true;
  } catch {
    return false;
  }
}
