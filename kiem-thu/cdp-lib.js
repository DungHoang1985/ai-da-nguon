// Hàm dùng chung cho các bài thử chạy trong khung add-in qua kiem-thu/cdp.mjs (ghép trước tệp bài thử).
// Không đọc kho lưu của khung (localStorage, IndexedDB): ở đó có API key thật.
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const label = (b) => `${b.getAttribute("aria-label") || ""} ${b.innerText}`.trim();
const btn = (re) => [...document.querySelectorAll("button")].find((b) => re.test(label(b)));
const setNative = (el, value) => {
  const proto = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : el.tagName === "SELECT" ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto, "value").set.call(el, value);
  el.dispatchEvent(new Event(el.tagName === "SELECT" ? "change" : "input", { bubbles: true }));
};
/** Về khung chat nếu đang ở trang Cài đặt */
const toChat = async () => {
  const back = btn(/Quay lại/);
  if (back) {
    back.click();
    await sleep(800);
  }
};
// Chỉ nút Dừng của lượt đang chạy (aria-label "Dừng"); nút "Dừng ở đây" của thẻ "Làm tiếp?" không tính là bận (24/11/2026)
const busy = () => !!btn(/^Dừng$/);
/** Chữ đang hiện trong khung (phần cuối) */
const tail = (n = 1800) => document.body.innerText.slice(-n);
/** Chọn nhà cung cấp và model (model: biểu thức khớp mã model, ví dụ /haiku-5-5/); reload = tải lại danh mục */
async function pick(providerId, modelRe, reload = false) {
  await toChat();
  // Nút "Model đang dùng" bật / tắt bảng chọn: bảng đang mở sẵn (lần trước bỏ dở) thì không bấm nữa
  if (!document.querySelector('select[aria-label="Nhà cung cấp"]')) {
    btn(/^Model đang dùng/).click();
    await sleep(700);
  }
  const sel = document.querySelector('select[aria-label="Nhà cung cấp"]');
  if (sel.value !== providerId) {
    setNative(sel, providerId);
    await sleep(1500);
  }
  if (reload) {
    btn(/Tải lại danh sách model/).click();
    await sleep(4000);
  }
  for (let i = 0; i < 20 && !document.querySelector('[role="option"]'); i++) await sleep(500);
  const opts = [...document.querySelectorAll('[role="listbox"][aria-label="Model"] [role="option"]')];
  const hit = opts.find((o) => modelRe.test(o.getAttribute("title") || ""));
  if (!hit) return { loi: "Không thấy model khớp", coSan: opts.map((o) => o.getAttribute("title")).slice(0, 40) };
  hit.click();
  await sleep(600);
  return { daChon: label(btn(/^Model đang dùng/)) };
}
/** Gửi một yêu cầu rồi chờ AI chạy xong. opts.approve(text) trả "allow" | "all" | "deny" cho từng lần hỏi (mặc định "all" nếu có nút, không thì "allow");
 *  opts.stopWhen() trả true thì bấm Dừng. Trả về các lần hỏi đã gặp và phần cuối khung. */
async function send(text, opts = {}) {
  await toChat();
  const ta = document.querySelector("textarea");
  ta.focus();
  setNative(ta, text);
  await sleep(300);
  const go = btn(/^Gửi/);
  if (!go || go.disabled) return { loi: "Nút Gửi không bấm được", khung: tail(600) };
  go.click();
  const asked = [];
  let stopped = false;
  const t0 = Date.now();
  await sleep(600);
  while (busy() && Date.now() - t0 < (opts.timeout || 300000)) {
    const allow = btn(/^Cho phép$/);
    if (allow) {
      const box = allow.parentElement.parentElement;
      const q = box.innerText.replace(/\n+/g, " | ");
      const all = btn(/^Cho phép hết lần này$/);
      const want = opts.approve ? opts.approve(q) : all ? "all" : "allow";
      asked.push(`${q.slice(0, 260)} → ${want}`);
      (want === "deny" ? btn(/^Bỏ qua$/) : want === "all" && all ? all : allow).click();
      await sleep(500);
    }
    if (!stopped && opts.stopWhen && opts.stopWhen()) {
      btn(/^Dừng$/)?.click();
      stopped = true;
    }
    await sleep(300);
  }
  await sleep(1000);
  return { daHoi: asked, daDung: stopped, giay: Math.round((Date.now() - t0) / 1000), khung: tail(opts.tail || 1500) };
}
