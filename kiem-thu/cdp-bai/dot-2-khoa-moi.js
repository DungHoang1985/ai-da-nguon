// Đợt 2: sau khi anh đăng nhập lại, kiểm khóa đang dùng có phải khóa MỚI không mà không đọc kho lưu và không ghi con số nào:
// khóa vừa được cấp thì mức "đã dùng" hôm nay và tháng này của chính khóa đó bằng 0. Chỉ trả về đúng/sai.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-khoa-moi.js
if (!document.querySelector('[data-aidn-card="signin-openrouter"]')) {
  btn(/^Cài đặt/)?.click();
  await sleep(1200);
}
const card = document.querySelector('[data-aidn-card="signin-openrouter"]');
if (!card) return { loi: "Không thấy thẻ OpenRouter" };
const inCard = (re) => [...card.querySelectorAll("button")].find((b) => re.test(b.innerText.trim()));
inCard(/^Làm mới$/)?.click();
await sleep(3500);
const t = (card.querySelector('[data-aidn="han-muc"]') || {}).innerText || "";
const m = /hôm nay ([\d.,]+) USD · tháng này ([\d.,]+) USD/.exec(t);
const zero = (s) => Number(String(s).replace(",", ".")) === 0;
const res = {
  coDongDaDung: !!m,
  daDungHomNayBangKhong: m ? zero(m[1]) : null,
  daDungThangNayBangKhong: m ? zero(m[2]) : null,
  dongTrangThai: [...card.querySelectorAll('[role="status"]')].map((el) => el.innerText.replace(/\d+/g, "#")),
};
await toChat();
return res;
