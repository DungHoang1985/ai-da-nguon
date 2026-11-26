// Bấm "Kiểm tra từng nguồn" (Cài đặt → Kiểm tra nhà cung cấp) rồi chờ xong; trả phiên bản và từng dòng kết quả (✅/❌ nguồn · model — lời).
// Gọi thật mỗi nguồn có key một lượt nhỏ (tốn rất ít token). Dùng để so hai bản (ví dụ trước / sau B9). Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/kiem-tra-tung-nguon.js
for (let i = 0; i < 60 && !btn(/^Cài đặt/); i++) await sleep(300);
const phienBan = (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0]; // khung hiện "v4.14" khi bản x.y.0
btn(/^Cài đặt/).click();
for (let i = 0; i < 30 && !btn(/^Kiểm tra từng nguồn$/); i++) await sleep(300);
const go = btn(/^Kiểm tra từng nguồn$/);
if (!go) return { phienBan, loi: "Không thấy nút Kiểm tra từng nguồn" };
if (go.disabled) return { phienBan, loi: "Nút bị tắt (chưa có nguồn nào có key)" };
go.scrollIntoView();
const t0 = Date.now();
go.click();
await sleep(1500);
while (Date.now() - t0 < 300000 && (btn(/^Dừng$/) || document.body.innerText.includes("⏳"))) await sleep(1000);
const dong = document.body.innerText.split("\n").filter((l) => /^(✅|❌) /.test(l.trim()) && / · | — /.test(l));
await toChat();
return { phienBan, giay: Math.round((Date.now() - t0) / 1000), dong };
