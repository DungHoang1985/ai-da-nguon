// Đợt 2: lối dự phòng dán mã, thử trong Office tới ngay trước bước anh phải đăng nhập.
//  - Ô "Dán mã vào đây": dán chuỗi sai dạng (bị từ chối tại chỗ) rồi dán một MÃ GIẢ hợp dạng (máy chủ OpenRouter từ chối).
//  - Nút "Mở trang lấy mã trong trình duyệt": MẶC ĐỊNH KHÔNG BẤM, vì bấm là mở thật một thẻ trong trình duyệt mặc định của máy
//    (Office.context.ui.openBrowserWindow không ghi đè được: writable = false, đo 16/11/2026). Muốn thử cả nút này thì chạy trước
//    một tệp đặt window.__aidnMoTrinhDuyet = true, và nói với anh là sẽ có một thẻ trình duyệt mở ra.
// Khóa đang lưu không đổi: chỉ đổi khi đổi mã thành công. Không đọc kho lưu.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-dan-ma.js
if (!document.querySelector('[data-aidn-card="signin-openrouter"]')) {
  btn(/^Cài đặt/)?.click();
  await sleep(1200);
}
const card = document.querySelector('[data-aidn-card="signin-openrouter"]');
if (!card) return { loi: "Không thấy thẻ OpenRouter", khung: tail(400) };
card.scrollIntoView();
const inCard = (re) => [...card.querySelectorAll("button")].find((b) => re.test(b.innerText.trim()));
if (!inCard(/^Mở trang lấy mã trong trình duyệt$/)) {
  inCard(/^Đăng nhập lại \(đổi khóa\)$/)?.click();
  await sleep(600);
}
const status = () => [...card.querySelectorAll('[role="status"]')].map((el) => el.innerText).join(" | ");
const out = {};

if (window.__aidnMoTrinhDuyet) {
  delete window.__aidnMoTrinhDuyet;
  const b = inCard(/^Mở trang lấy mã trong trình duyệt$/);
  for (let i = 0; i < 20 && b.disabled; i++) await sleep(300);
  b.click();
  await sleep(800);
  out.b1_sauKhiBamMoTrinhDuyet = status();
}

const box = card.querySelector('input[aria-label="Dán mã vào đây"]');
const confirm = () => inCard(/^Xác nhận mã$/);
out.b2_nutTatKhiOTrong = confirm().disabled;
setNative(box, "ma co dau cach");
await sleep(300);
out.b2_nutBatKhiCoChu = !confirm().disabled;
confirm().click();
await sleep(600);
out.b2_maCoDauCach = status();
setNative(box, "https://openrouter.ai/auth");
await sleep(300);
confirm().click();
await sleep(600);
out.b2_diaChiKhongCoMa = status();
// Dán nhầm một chuỗi có dạng API key của nhà cung cấp khác (chuỗi giả): phải bị từ chối tại chỗ, không gửi đi đâu
let goiMang = 0;
const realFetch = window.fetch;
window.fetch = (...a) => {
  goiMang++;
  return realFetch(...a);
};
try {
  setNative(box, "AIza" + "x".repeat(35));
  await sleep(300);
  confirm().click();
  await sleep(700);
} finally {
  window.fetch = realFetch;
}
out.b2_danNhamKeyNguonKhac = status();
out.b2_soLanGoiMangKhiDanNhamKey = goiMang;
setNative(box, "ma-gia-0123456789abcdef");
await sleep(300);
confirm().click();
await sleep(400);
for (let i = 0; i < 50 && /Đang đổi mã/.test(status()); i++) await sleep(300);
await sleep(500);
out.b3_maGiaHopDang = status();
out.b3_oNhapDuocXoa = box.value === "";
out.sauCung = { conDangNhap: /Hạn mức/.test(card.innerText), nut: [...card.querySelectorAll("button")].map((b) => b.innerText.trim()) };
await toChat();
out.modelVanCon = btn(/^Model đang dùng/) ? label(btn(/^Model đang dùng/)).slice(0, 60) : "(không thấy nhãn model)";
return out;
