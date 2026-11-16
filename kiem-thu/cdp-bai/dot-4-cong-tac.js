// Đợt 4: đặt hai công tắc "Dấu vết soát lại (Excel)" trong trang Cài đặt bằng cú bấm thật, trả về trạng thái TRƯỚC khi đổi (để trả lại sau khi thử).
// Tham số ở tệp ghép phía trước: window.__aidnCongTac = { nhatKy: true|false, toSang: true|false } (thiếu khóa nào thì không đụng công tắc đó)
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-4-cong-tac.js
const want = window.__aidnCongTac || {};
delete window.__aidnCongTac;
if (!btn(/Chạy tự kiểm tra/)) {
  btn(/Cài đặt/)?.click();
  await sleep(1200);
}
const sw = (re) => [...document.querySelectorAll('input[role="switch"], input[type="checkbox"]')].find((el) => re.test(el.closest("label, div")?.innerText || el.getAttribute("aria-label") || el.parentElement?.parentElement?.innerText || ""));
const pairs = [
  ["nhatKy", /Ghi nhật ký vào sheet/],
  ["toSang", /Tô nền vàng nhạt/],
];
const truoc = {};
const sau = {};
for (const [key, re] of pairs) {
  const el = sw(re);
  if (!el) return { loi: `Không thấy công tắc ${key}`, khung: tail(800) };
  truoc[key] = el.checked;
  if (want[key] !== undefined && el.checked !== want[key]) {
    el.scrollIntoView();
    el.click();
    await sleep(500);
  }
  sau[key] = sw(re).checked;
}
const canhBao = /đi theo tệp khi anh gửi cho người khác/.test(document.body.innerText);
await toChat();
return { truoc, sau, canhBaoKhiBat: canhBao };
