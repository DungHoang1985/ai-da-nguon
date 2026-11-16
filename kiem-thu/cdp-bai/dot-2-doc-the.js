// Đợt 2: đọc dòng trạng thái và các nút của thẻ OpenRouter đang hiện (không bấm gì, không đọc kho lưu). Số được che bằng #.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-doc-the.js
const che = (s) => s.replace(/\d+(?:[.,]\d+)*/g, "#");
const card = document.querySelector('[data-aidn-card="signin-openrouter"]') || document.querySelector('[data-aidn-card="welcome-openrouter"]');
if (!card) return { loi: "Không thấy thẻ OpenRouter (đang ở khung chat?)", khung: che(tail(300)) };
return {
  trangThai: [...card.querySelectorAll('[role="status"]')].map((el) => che(el.innerText)),
  conDangNhap: /Đã đăng nhập|Đã có khóa|Hạn mức/.test(card.innerText),
  coTheHanMuc: !!card.querySelector('[data-aidn="han-muc"]'),
  nut: [...card.querySelectorAll("button")].map((b) => b.innerText.trim()),
};
