// Thăm dò Đợt 1A (bước 4): ba điểm lần soát mã độc lập ghi "chưa chắc", kiểm trong Word thật.
//  (a) Định dạng chữ khi đang TrackAll có sinh sửa đổi được theo dõi không; đổi cấp đánh số thì sao.
//  (b) Tìm trong đoạn bằng chuỗi có khoảng trắng ở cuối / ở đầu (neo để chèn thuần) có ra đúng số lần không.
//  (c) table.values có lẫn chữ đã xóa có theo dõi không.
// Chạy trên hợp đồng thử: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-4-sau-soat-ma.js  (đóng tệp không lưu sau khi chạy)
const out = {};
const test = async (name, fn) => {
  try {
    out[name] = await fn();
  } catch (e) {
    out[name] = { LOI: (e && e.message) || String(e), o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
};
const count = async (ctx) => {
  const tc = ctx.document.body.getTrackedChanges();
  tc.load("items/type,items/text");
  await ctx.sync();
  return tc.items;
};
await test("a. định dạng và đổi cấp khi TrackAll", () =>
  Word.run(async (ctx) => {
    ctx.document.changeTrackingMode = "TrackAll";
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const i = ps.items.findIndex((p) => p.text.startsWith("Hợp đồng có hiệu lực từ ngày ký"));
    const n0 = (await count(ctx)).length;
    const hit = ps.items[i].search("hiệu lực", { matchCase: true });
    hit.load("items");
    await ctx.sync();
    hit.items[0].font.bold = true;
    await ctx.sync();
    const afterFormat = await count(ctx);
    const li = ps.items[i].listItemOrNullObject;
    li.load("level,listString");
    await ctx.sync();
    const lv = li.level;
    li.level = lv + 1;
    await ctx.sync();
    const afterLevel = await count(ctx);
    li.level = lv;
    hit.items[0].font.bold = false;
    await ctx.sync();
    ctx.document.changeTrackingMode = "Off";
    await ctx.sync();
    return {
      suaDoiTruoc: n0,
      sauDinhDang: afterFormat.length,
      loaiMoi: afterFormat.slice(n0 > 0 ? 0 : 0).filter((t) => t.type === "Formatted").map((t) => `${t.type} ${JSON.stringify(t.text.slice(0, 30))}`),
      sauDoiCap: afterLevel.length,
    };
  })
);
await test("b. tìm bằng chuỗi có khoảng trắng đầu / cuối", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const p = ps.items.find((x) => x.text.startsWith("Hợp đồng có hiệu lực từ ngày ký"));
    const t = p.text;
    const q = ["lực ", " từ", "ngày ", " ", "  "];
    const hs = q.map((s) => {
      const h = p.search(s, { matchCase: true });
      h.load("items/text");
      return h;
    });
    await ctx.sync();
    const occ = (s) => {
      let n = 0;
      for (let pos = 0; ; n++) {
        const k = t.indexOf(s, pos);
        if (k < 0) break;
        pos = k + s.length;
      }
      return n;
    };
    return Object.fromEntries(q.map((s, k) => [JSON.stringify(s), { wordTimRa: hs[k].items.length, tinhRa: occ(s), chuKetQuaDau: hs[k].items[0] ? JSON.stringify(hs[k].items[0].text) : null }]));
  })
);
await test("c. table.values khi ô có chữ đã xóa có theo dõi", () =>
  Word.run(async (ctx) => {
    ctx.document.changeTrackingMode = "Off";
    const last = ctx.document.body.paragraphs.getLast();
    const tb = last.insertTable(2, 2, "Before", [
      ["Hạng mục thử", "Số tiền"],
      ["Tiền cọc", "24 tháng"],
    ]);
    await ctx.sync();
    ctx.document.changeTrackingMode = "TrackAll";
    const cell = tb.getCell(1, 1);
    const h = cell.body.search("24", { matchCase: true });
    h.load("items");
    await ctx.sync();
    h.items[0].insertText("12", "Replace");
    await ctx.sync();
    tb.load("values");
    const cp = cell.body.paragraphs.getFirst();
    cp.load("text");
    const cur = cp.getText();
    await ctx.sync();
    const res = { "table.values[1][1]": tb.values[1][1], "paragraph.text": cp.text, "getText()": cur.value };
    ctx.document.changeTrackingMode = "Off";
    tb.delete();
    await ctx.sync();
    return res;
  })
);
return out;
