// Đợt 2: nhớ nhà cung cấp đang chọn vào window.__aidnNguonCu (để bài dọn dẹp trả lại), không đổi gì.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-2-nho-nguon.js
await toChat();
const b = btn(/^Model đang dùng/);
if (!b) return { loi: "Không thấy nhãn model (màn chào?)", khung: tail(300) };
b.click();
await sleep(800);
const sel = document.querySelector('select[aria-label="Nhà cung cấp"]');
if (!window.__aidnNguonCu) window.__aidnNguonCu = sel.value;
const res = { nguonDangChon: sel.value, nguonCuDaNho: window.__aidnNguonCu, cacNguonCoKey: [...sel.options].map((o) => o.value) };
document.body.click();
await sleep(300);
return res;
