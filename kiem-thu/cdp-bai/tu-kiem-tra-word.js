// Bấm "Chạy tự kiểm tra" ở Cài đặt (Word) rồi chờ xong; trả phiên bản, tổng kết và từng bước. Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/tu-kiem-tra-word.js
for (let i = 0; i < 60 && !btn(/^Cài đặt/); i++) await sleep(300);
const phienBan = (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0]; // khung hiện "v4.14" khi bản x.y.0
btn(/^Cài đặt/).click();
for (let i = 0; i < 30 && !btn(/^Chạy tự kiểm tra$/); i++) await sleep(300);
const t0 = Date.now();
btn(/^Chạy tự kiểm tra$/).click();
await sleep(1500);
while (Date.now() - t0 < 540000 && !/bước đạt/.test(document.body.innerText)) await sleep(1000);
const txt = document.body.innerText;
const tongKet = (txt.match(/(Tất cả đạt|\d+ lỗi) · \d+\/\d+ bước đạt/) || [""])[0];
const buoc = txt.split("\n").filter((l) => /^(✅|❌|⚠️)/.test(l.trim()));
await toChat();
return { phienBan, giay: Math.round((Date.now() - t0) / 1000), tongKet, buoc };
