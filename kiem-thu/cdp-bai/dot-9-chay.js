// Đợt 9: Chạy thật một bài Word. Biến: NGUON, MODEL_RE, YEU_CAU, (tùy chọn) TREO_NGUON = tên miền giả lập im lặng (chỉ trong bài thử, không đổi mã)
if (typeof TREO_NGUON === "string") {
  if (!window.__aidnFetch) window.__aidnFetch = window.fetch;
  const orig = window.__aidnFetch;
  window.fetch = (u, init) => {
    const url = String(u && u.url ? u.url : u);
    if (url.includes(TREO_NGUON)) {
      window.__aidnTreo = (window.__aidnTreo || 0) + 1;
      return new Promise((_res, rej) => init && init.signal && init.signal.addEventListener("abort", () => rej(new DOMException("aborted", "AbortError"))));
    }
    return orig(u, init);
  };
}
// Bảng chọn đang mở (bài trước để lại) thì đóng trước, không thì pick() bấm vào lại đóng mất
if (document.querySelector('select[aria-label="Nhà cung cấp"]')) { btn(/^Model đang dùng/).click(); await sleep(500); }
const chon = await pick(NGUON, MODEL_RE);
const r = await send(YEU_CAU, { timeout: 400000, tail: 2500 });
if (window.__aidnFetch) {
  window.fetch = window.__aidnFetch;
  delete window.__aidnFetch;
}
const last = [...document.querySelectorAll("*")].length;
void last;
return {
  model: chon,
  treo: window.__aidnTreo || 0,
  ...r,
  doiNguon: (r.khung.match(/↪[^\n]+/g) || []),
  lanGoiAI: ((r.khung.match(/(\d+) lần gọi AI/) || [])[1]) || null,
};
