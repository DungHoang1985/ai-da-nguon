// Đợt 3 (Chạy thật, tiêu chí B-1, B-4 và kiểm bài): giao việc cho AI trong PowerPoint rồi ĐỌC LẠI bài bằng mã.
// Tham số ở tệp ghép phía trước: window.__aidnBai = { yeuCau, nguon?, model?, anh?: [số slide cần chụp trước / sau] }
// Trả: các bước công cụ hiện trên khung, số slide trước / sau, chữ và khung của từng slide mới hoặc slide được nêu, ảnh slide (base64 rút gọn độ dài).
// Chạy: node kiem-thu/cdp.mjs eval 9229/Powerpoint kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-3-ppt-bai.js
const bai = window.__aidnBai || {};
delete window.__aidnBai;
if (!bai.yeuCau) return { loi: "Thiếu window.__aidnBai.yeuCau" };

const doc = (only) =>
  PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const out = [];
    for (let i = 0; i < slides.items.length; i++) {
      if (only && !only.includes(i + 1)) continue;
      const s = slides.items[i];
      s.layout.load("name");
      s.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      await ctx.sync();
      const shapes = [];
      for (const sh of s.shapes.items) {
        const o = { name: sh.name, type: String(sh.type), box: [sh.left, sh.top, sh.width, sh.height].map((v) => Math.round(v)) };
        if (["TextBox", "GeometricShape", "Placeholder"].includes(o.type)) {
          try {
            const tr = sh.textFrame.textRange;
            tr.load("text");
            tr.font.load("size,name");
            sh.textFrame.load("autoSizeSetting");
            await ctx.sync();
            o.text = tr.text.slice(0, 400);
            o.size = tr.font.size;
            o.font = tr.font.name;
            o.auto = String(sh.textFrame.autoSizeSetting);
            // Ước lượng tràn giống phép đo của add-in: mỗi ký tự rộng nửa cỡ chữ, dòng cao 1,2 cỡ chữ
            if (o.text.trim() && o.size && o.auto === "AutoSizeNone") {
              const perLine = Math.max(1, Math.floor((sh.width - 14.4) / (o.size * 0.5)));
              const need = o.text.split(/\r\n|\r|\n|\v/).reduce((h, l) => h + Math.max(1, Math.ceil(l.length / perLine)) * o.size * 1.2, 7.2);
              o.uocCao = Math.round(need);
              o.tran = need > sh.height * 1.15 + 4;
            }
          } catch (e) {
            /* không có chữ */
          }
        }
        if (o.type === "Table") {
          try {
            const t = sh.getTable();
            t.load("values");
            await ctx.sync();
            o.table = t.values;
          } catch (e) {
            /* bỏ qua */
          }
        }
        shapes.push(o);
      }
      out.push({ slide: i + 1, layout: s.layout.name, shapes });
    }
    return { soSlide: slides.items.length, slides: out };
  });
const chup = (n) =>
  PowerPoint.run(async (ctx) => {
    const img = ctx.presentation.slides.getItemAt(n - 1).getImageAsBase64({ width: 800 });
    await ctx.sync();
    return img.value;
  }).catch((e) => "LOI " + e.message);

await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
if (bai.nguon) {
  const pk = await pick(bai.nguon, new RegExp(bai.model, "i"));
  document.body.click();
  await sleep(300);
  if (pk.loi) return { loi: "Không chọn được model", chiTiet: pk };
}
const truoc = await doc(bai.xem);
const anhTruoc = {};
for (const n of bai.anh || []) anhTruoc[n] = await chup(n);
const tong0 = (await doc([0])).soSlide;
const r = await send(bai.yeuCau, { timeout: 480000, tail: 3000 });
const tong1 = (await doc([0])).soSlide;
const moi = [];
for (let n = tong0 + 1; n <= tong1; n++) moi.push(n);
const sau = await doc([...(bai.xem || []), ...moi]);
const anhSau = {};
for (const n of [...(bai.anh || []), ...moi.slice(0, 4)]) anhSau[n] = await chup(n);
const khung = document.body.innerText;
window.__aidnAnh = { truoc: anhTruoc, sau: anhSau };
return {
  model: label(btn(/^Model đang dùng/)).slice(-40),
  ketNoi: document.querySelector("[data-links]")?.innerText || null,
  yeuCau: bai.yeuCau,
  giay: r.giay,
  // Số lần gọi model của nhiệm vụ (khung ghi ở thuộc tính data-calls của dòng token) và dòng hiện trên khung
  lanGoiAI: Number([...document.querySelectorAll("[data-calls]")].pop()?.getAttribute("data-calls") || 0),
  dongToken: ([...document.querySelectorAll("[data-calls]")].pop()?.innerText || "").slice(0, 160),
  daHoi: r.daHoi.map((q) => q.slice(0, 160)),
  buocCongCu: (khung.match(/^[✅⚠️❌][^\n]*/gm) || []).map((x) => x.slice(0, 170)),
  soSlideTruoc: tong0,
  soSlideSau: tong1,
  truoc: truoc.slides,
  sau: sau.slides,
  anh: { truoc: Object.keys(anhTruoc).map((k) => `slide ${k}: ${anhTruoc[k].length} ký tự`), sau: Object.keys(anhSau).map((k) => `slide ${k}: ${anhSau[k].length} ký tự`) },
  canhBaoNoiDaSua: !!document.querySelector("[data-claim-warning]"),
  cuoiKhung: r.khung.slice(-1500),
};
