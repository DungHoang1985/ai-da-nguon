// 4.12.0 (Chạy thật): cổng cuối lượt "đổi hình thức mà chưa xem lại" ở PowerPoint.
// Giao việc thêm một slide và dặn trả lời "Xong." không cần kiểm → cổng phải nhắc, AI gọi chụp ảnh / kiểm bài rồi mới xong.
// Xong thì xóa slide thử. Đổi nguồn / model bằng window.__aidnCb = { nguon, model }. Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/cong-xem-lai-ppt.js
const dem = () => PowerPoint.run(async (ctx) => { const s = ctx.presentation.slides; s.load("items"); await ctx.sync(); return s.items.length; });
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const cb = window.__aidnCb || {};
delete window.__aidnCb;
const pk = await pick(cb.nguon || "openrouter", new RegExp(cb.model || "nemotron-3-ultra", "i"));
document.body.click();
await sleep(300);
if (pk.loi) return { loi: "Không chọn được model", chiTiet: pk };
const truoc = await dem();
const tag = "XL" + Date.now().toString(36).slice(-4).toUpperCase();
const r = await send(`Thêm MỘT slide mới ở cuối bài, tiêu đề “Kiểm tra cổng ${tag}”, có một ô chữ ghi “Nội dung thử”. Làm xong trả lời đúng một câu “Xong.”, không cần kiểm tra lại.`, { tail: 2500 });
const sau = await dem();
const buoc = (document.body.innerText.match(/[✅❌⚠️] [^\n]{0,120}/g) || []).slice(-10);
const k = r.khung;
const iXong = k.indexOf("Xong.");
const xemSauXong = iXong >= 0 && /Chụp|ảnh slide|Kiểm (tra )?bài|check/i.test(k.slice(iXong));
// Dọn: xóa các slide thêm mới
await PowerPoint.run(async (ctx) => {
  const s = ctx.presentation.slides;
  s.load("items");
  await ctx.sync();
  for (let i = s.items.length - 1; i >= truoc; i--) s.items[i].delete();
  await ctx.sync();
});
return { model: pk.daChon, slideThem: sau - truoc, xemSauXong, buoc, giay: r.giay, khung: k.slice(-1500) };
