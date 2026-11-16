// Đợt 2: chạy TRONG hộp thoại đăng nhập (trang của nhà cung cấp) để xem nó đang ở đâu và mang tham số gì.
// Chỉ trả tên miền, đường dẫn, tên tham số; giá trị chỉ in với tham số không nhạy cảm, còn lại in độ dài. Không đọc cookie, không đọc kho lưu.
// Chạy: node kiem-thu/cdp.mjs eval "9229#openrouter.ai" kiem-thu/cdp-bai/dot-2-doc-hop-thoai.js
const SHOW = new Set(["callback_url", "code_challenge_method", "key_label"]);
const mask = (u) => [...u.searchParams.entries()].map(([k, v]) => (SHOW.has(k) ? `${k}=${v}` : `${k}=<${v.length} ký tự>`));
const here = new URL(location.href);
const out = { tenMien: here.hostname, duongDan: here.pathname, thamSo: mask(here), tieuDe: document.title.slice(0, 80) };
// OpenRouter chuyển người chưa đăng nhập sang trang đăng nhập, giữ địa chỉ gốc trong redirect_url
const back = here.searchParams.get("redirect_url");
if (back) {
  try {
    const b = new URL(back, here.origin);
    out.diaChiGoc = { tenMien: b.hostname, duongDan: b.pathname, thamSo: mask(b) };
    const c = b.searchParams.get("code_challenge") || "";
    out.maThachThucDungDang = /^[A-Za-z0-9_-]{43}$/.test(c);
  } catch (e) {
    out.diaChiGoc = "không đọc được";
  }
} else {
  const c = here.searchParams.get("code_challenge") || "";
  if (c) out.maThachThucDungDang = /^[A-Za-z0-9_-]{43}$/.test(c);
}
out.chuTrenTrang = (document.body.innerText || "").replace(/\s+/g, " ").slice(0, 300);
out.coOnhapEmail = !!document.querySelector('input[type="email"], input[name="identifier"], input[name="emailAddress"]');
out.nutDangNhap = [...document.querySelectorAll("button, a")].map((b) => (b.innerText || "").trim()).filter((t) => /google|github|continue|sign|metamask|đăng/i.test(t)).slice(0, 10);
return out;
