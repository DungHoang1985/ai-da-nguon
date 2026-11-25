// 4.12.0 (Chạy thật): khối "đoạn người dùng tự sửa giữa hai lượt". Lượt 1 cho AI đọc; sửa tay một đoạn (không qua AI);
// lượt 2 hỏi AI người dùng vừa sửa gì → câu trả lời phải nêu đúng chữ mới. Dọn văn bản thử khi xong. Không đọc kho lưu của khung.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-lib.js kiem-thu/cdp-bai/doan-tu-sua.js
await Word.run(async (ctx) => {
  ctx.document.body.clear();
  ctx.document.body.insertParagraph("Điều 4. Phạt vi phạm (văn bản thử đoạn tự sửa)", "Start");
  ctx.document.body.insertParagraph("Mức phạt vi phạm là 5% giá trị phần nghĩa vụ bị vi phạm.", "End");
  await ctx.sync();
});
await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(600);
const cb = window.__aidnCb || {};
delete window.__aidnCb;
const pk = await pick(cb.nguon || "openrouter", new RegExp(cb.model || "gemma-4-31b", "i"));
document.body.click();
await sleep(300);
if (pk.loi) return { loi: "Không chọn được model", chiTiet: pk };
const r1 = await send("Đọc văn bản đang mở rồi chỉ trả lời “Đã đọc.”", { tail: 400 });
await Word.run(async (ctx) => {
  ctx.document.changeTrackingMode = "Off";
  const ps = ctx.document.body.paragraphs;
  ps.load("items");
  await ctx.sync();
  const h = ps.items[1].search("5%", { matchCase: true });
  h.load("items");
  await ctx.sync();
  h.items[0].insertText("8%", "Replace");
  await ctx.sync();
});
const r2 = await send("Kể từ lượt trước tôi đã tự sửa chỗ nào trong văn bản? Trả lời ngắn, không gọi công cụ sửa.", { tail: 1200 });
const traLoi = r2.khung;
await Word.run(async (ctx) => {
  ctx.document.body.clear();
  await ctx.sync();
});
return { model: pk.daChon, luot1: r1.giay, nhacDung8: /8\s?%/.test(traLoi.slice(traLoi.lastIndexOf("Kể từ lượt trước"))), khung: traLoi.slice(-900) };
