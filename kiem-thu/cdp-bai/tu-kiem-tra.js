// Bấm nút "Chạy tự kiểm tra" trong trang Cài đặt rồi chờ xong, trả về từng bước và dòng tổng kết.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/tu-kiem-tra.js
const version = (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0];
if (!btn(/Chạy tự kiểm tra/)) {
  const gear = btn(/Cài đặt/);
  if (!gear) return { loi: "Không thấy nút Cài đặt", khung: tail(400) };
  gear.click();
  await sleep(1200);
}
const go = btn(/Chạy tự kiểm tra/);
if (!go) return { loi: "Không thấy nút Chạy tự kiểm tra", khung: tail(600) };
go.scrollIntoView();
go.click();
const t0 = Date.now();
await sleep(1500);
while (btn(/Chạy tự kiểm tra/)?.disabled && Date.now() - t0 < 540000) await sleep(1000);
await sleep(500);
const lines = [...document.querySelectorAll("span, div")]
  .filter((el) => el.children.length === 0 && /^[✅❌➖] /.test(el.innerText || ""))
  .map((el) => el.innerText);
const tong = [...document.querySelectorAll("span, div")].map((el) => (el.children.length === 0 ? el.innerText : "")).find((t) => /\d+\/\d+ bước đạt/.test(t)) || "";
await toChat();
return { phienBan: version, giay: Math.round((Date.now() - t0) / 1000), tongKet: tong, buoc: lines };
