// Đợt 1B: tiêu đề các hội thoại do bài Chạy thật tạo ra, và nguồn phải trả lại. Ghép trước don-dep-sau-khi-thu.js:
//   node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/word-1b-bai.js kiem-thu/cdp-bai/word-1b-don-dep.js kiem-thu/cdp-bai/don-dep-sau-khi-thu.js
// Trước đó tắt hai tùy chọn thử: window.__aidnCongTac = { nghiKy: false, soatLapLuan: false } + word-1b-cong-tac.js
const escRe1b = (s) => "^" + s.slice(0, 40).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
window.__aidnTieuDeThem = Object.values(window.__aidnBai1B || {}).map(escRe1b);
window.__aidnNguonCu = "openrouter";
