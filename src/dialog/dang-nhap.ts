/* global Office */
/** Trang mở trong hộp thoại của Office khi bấm "Đăng nhập" (Đợt 2).
 *  - Lần mở đầu (?p=<nguồn>&c=<mã thách thức>&s=<state>): tự chuyển tới trang đăng nhập của nhà cung cấp.
 *    Tên miền đích lấy từ cấu hình trong mã, không lấy từ tham số (không thành chỗ chuyển hướng tùy ý).
 *  - Lần nhà cung cấp chuyển về (?code=…&state=…): gửi mã về khung add-in bằng messageParent rồi gỡ mã khỏi thanh địa chỉ.
 *    Không gửi được (trang mở ngoài Office) thì hiện mã để chép và dán vào add-in.
 *  Trang này không giữ bí mật nào: mã kiểm (code_verifier) nằm trong khung add-in, thiếu nó thì mã vô dụng. */
import { buildAuthUrl, dialogMessage, parseDialogQuery } from "../core/oauth";
import { getProvider } from "../core/providers";

const el = (id: string) => document.getElementById(id) as HTMLElement;

function show(text: string, o: { err?: boolean; title?: string } = {}): void {
  if (o.title) el("title").textContent = o.title;
  const t = el("text");
  t.textContent = text;
  t.className = o.err ? "err" : "";
}

function showCode(code: string): void {
  show("Chép mã dưới đây, quay lại add-in AI Đa Nguồn, dán vào ô “Dán mã vào đây” rồi bấm “Xác nhận mã”. Mã chỉ dùng một lần, trong 10 phút.", {
    title: "Mã đăng nhập",
  });
  const input = el("code") as HTMLInputElement;
  input.value = code;
  el("codeBox").style.display = "block";
  el("copy").onclick = async () => {
    try {
      await navigator.clipboard.writeText(code);
    } catch {
      input.select();
      document.execCommand("copy");
    }
    el("copied").textContent = " Đã chép.";
  };
}

/** Chờ office.js khởi động (tối đa `ms`). Trả true nếu đang chạy trong Office. */
function officeReady(ms: number): Promise<boolean> {
  return new Promise((resolve) => {
    const timer = setTimeout(() => resolve(false), ms);
    try {
      if (typeof Office === "undefined" || !Office.onReady) {
        clearTimeout(timer);
        resolve(false);
        return;
      }
      Office.onReady((info) => {
        clearTimeout(timer);
        resolve(!!info?.host);
      });
    } catch {
      clearTimeout(timer);
      resolve(false);
    }
  });
}

/** Gỡ mã khỏi thanh địa chỉ (và khỏi lịch sử của cửa sổ) */
function scrubAddress(): void {
  try {
    const rs = (window as any).__aidnReplaceState;
    if (typeof rs === "function") rs(null, "", window.location.pathname);
  } catch {
    /* không gỡ được: mã đã dùng một lần nên cũng hết giá trị */
  }
}

function sendToParent(message: string): boolean {
  try {
    Office.context.ui.messageParent(message);
    return true;
  } catch {
    return false;
  }
}

async function main(): Promise<void> {
  const q = parseDialogQuery(window.location.search);
  if (q.kind === "start") {
    const p = getProvider(q.providerId);
    if (!p?.oauth) {
      show("Nguồn này không hỗ trợ đăng nhập. Đóng cửa sổ này.", { err: true });
      return;
    }
    let url: string;
    try {
      url = buildAuthUrl(p.oauth, { challenge: q.challenge, state: q.state, callbackUrl: window.location.origin + window.location.pathname });
    } catch {
      show("Không dựng được địa chỉ đăng nhập. Đóng cửa sổ này và dùng lối dán mã trong add-in.", { err: true });
      return;
    }
    show(`Đang mở trang đăng nhập ${p.name}…`);
    // Chờ office.js khởi động xong rồi mới rời trang: nó ghi thông tin của Office vào kho phiên của cửa sổ,
    // nhờ đó khi nhà cung cấp chuyển về, trang này còn gửi được tin nhắn cho khung add-in.
    await officeReady(4000);
    window.location.replace(url);
    return;
  }
  if (q.kind === "callback") {
    show("Đang hoàn tất đăng nhập…");
    const inOffice = await officeReady(8000);
    const sent = inOffice && sendToParent(dialogMessage(q));
    scrubAddress();
    if (sent) {
      show("Đã gửi mã về add-in. Cửa sổ này sẽ tự đóng.");
      // Khung không đóng cửa sổ sau 6 giây (không nhận được tin nhắn): vẫn còn lối chép mã
      setTimeout(() => showCode(q.code), 6000);
    } else showCode(q.code);
    return;
  }
  if (q.kind === "denied") {
    const inOffice = await officeReady(4000);
    if (inOffice) sendToParent(JSON.stringify({ t: "aidn-oauth", denied: true }));
    show("Chưa được cho phép ở trang của nhà cung cấp. Đóng cửa sổ này rồi bấm đăng nhập lại nếu muốn.", { err: true });
    return;
  }
  show("Trang này chỉ dùng cho nút Đăng nhập trong add-in AI Đa Nguồn. Có thể đóng cửa sổ này.");
}

main();
