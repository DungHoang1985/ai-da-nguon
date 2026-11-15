// Thăm dò Đợt 1A (bước 2): sửa nhiều chỗ trong một đoạn khi đang theo dõi thay đổi, tìm kiếm trên đoạn có chữ đã xóa,
// chèn đoạn kế thừa đánh số, định danh đoạn, và dữ liệu cần cho Hoàn tác (tác giả, ngày của sửa đổi do add-in tạo).
// Chạy sau tham-do-1 trên cùng văn bản: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-2-sua.js
// Tên tác giả không phải "Đối tác" được che trước khi in (kho công khai).
const out = {};
const who = (a) => (a === "Đối tác" ? a : "(tài khoản Office đang dùng)");
const tcList = (tc) => tc.items.map((t) => `${t.type} ${JSON.stringify(t.text)} · ${who(t.author)} · ${new Date(t.date).toISOString()}`);
const test = async (name, fn) => {
  try {
    out[name] = await fn();
  } catch (e) {
    out[name] = { LOI: (e && e.message) || String(e), o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
};
const find = async (ctx, start) => {
  const ps = ctx.document.body.paragraphs;
  ps.load("items/text");
  await ctx.sync();
  const i = ps.items.findIndex((p) => p.text.startsWith(start));
  return { ps, i, p: ps.items[i] };
};
const setMode = async (ctx, mode) => {
  ctx.document.changeTrackingMode = mode;
  await ctx.sync();
};

// (a) Sửa 4 chỗ trong một đoạn, một lần sync, đang TrackAll: đổi "30" thứ hai (đang đậm), đổi "0,05%", chèn thuần, xóa thuần
await test("a. sửa 4 chỗ trong một đoạn (TrackAll)", () =>
  Word.run(async (ctx) => {
    const { p, i } = await find(ctx, "Bên B thanh toán trong 30 ngày");
    await setMode(ctx, "TrackAll");
    const h30 = p.search("30", { matchCase: true });
    const hPct = p.search("0,05%", { matchCase: true });
    const hDel = p.search(" trên số tiền chậm trả", { matchCase: true });
    const hAnchor = p.search("thanh toán", { matchCase: true });
    h30.load("items");
    hPct.load("items");
    hDel.load("items");
    hAnchor.load("items");
    await ctx.sync();
    const counts = { "30": h30.items.length, "0,05%": hPct.items.length, xoa: hDel.items.length, neo: hAnchor.items.length };
    // từ phải sang trái
    hDel.items[0].delete();
    hPct.items[0].insertText("0,03%", "Replace");
    h30.items[1].insertText("45", "Replace");
    hAnchor.items[0].insertText(" đủ", "After");
    await ctx.sync();
    const tc = p.getTrackedChanges();
    tc.load("items/type,items/author,items/text,items/date");
    const cur = p.getText();
    const bold = p.search("45 ngày", { matchCase: true });
    bold.load("items/font/bold,items/text");
    const first30 = p.search("trong 30 ngày", { matchCase: true });
    first30.load("items/font/bold");
    p.load("uniqueLocalId");
    await ctx.sync();
    await setMode(ctx, "Off");
    return {
      doan: i,
      soLanTim: counts,
      sauKhiSua: cur.value,
      suaDoiTrongDoan: tcList(tc),
      "45 ngày còn đậm": bold.items.map((r) => r.font.bold),
      "30 ngày đầu vẫn thường": first30.items.map((r) => r.font.bold),
    };
  })
);

// (b) Tìm kiếm trên đoạn có chữ đã xóa có theo dõi
await test("b. tìm trên đoạn có chữ đã xóa", () =>
  Word.run(async (ctx) => {
    const { p } = await find(ctx, "Thời hạn thuê là 24 tháng12 tháng");
    const q = ["24 tháng", "là 12 tháng", "12 tháng kể", "là 24 tháng", "tháng12", "tháng"];
    const hs = q.map((s) => {
      const h = p.search(s, { matchCase: true });
      h.load("items/text");
      return h;
    });
    const words = p.getTextRanges([" "], false);
    words.load("items/text");
    const whole = p.getRange("Content");
    let pos = null;
    try {
      whole.load("start,end");
      await ctx.sync();
      pos = [whole.start, whole.end];
    } catch (e) {
      pos = "Range.start/end: " + e.message;
    }
    await ctx.sync();
    return {
      soLanTim: Object.fromEntries(q.map((s, k) => [s, hs[k].items.length])),
      getTextRanges: words.items.map((w) => w.text),
      "Range.start/end của đoạn": pos,
    };
  })
);

// (b2) Sửa mức từ trên đoạn đang có sửa đổi của đối tác: đổi "ngày ký" → "ngày bàn giao", đổi "12 tháng" (chữ đối tác thêm) → "18 tháng"
await test("b2. sửa trên đoạn đang có sửa đổi của đối tác", () =>
  Word.run(async (ctx) => {
    const { p } = await find(ctx, "Thời hạn thuê là 24 tháng12 tháng");
    await setMode(ctx, "TrackAll");
    const h1 = p.search("ký", { matchCase: true });
    const h2 = p.search("12", { matchCase: true });
    h1.load("items");
    h2.load("items");
    await ctx.sync();
    const n = [h1.items.length, h2.items.length];
    h1.items[h1.items.length - 1].insertText("bàn giao", "Replace");
    h2.items[0].insertText("18", "Replace");
    await ctx.sync();
    const tc = p.getTrackedChanges();
    tc.load("items/type,items/author,items/text,items/date");
    const cur = p.getText();
    const orig = p.getReviewedText("Original");
    p.load("text");
    await ctx.sync();
    await setMode(ctx, "Off");
    return { soLanTim: n, banHienTai: cur.value, banGoc: orig.value, ".text": p.text, suaDoi: tcList(tc) };
  })
);

// (c) Định danh đoạn: giữ nguyên sau khi chèn đoạn phía trên và sau khi sửa chữ trong đoạn
await test("c. uniqueLocalId sau khi chèn phía trên và sửa trong đoạn", () =>
  Word.run(async (ctx) => {
    const a = await find(ctx, "Giá trị hợp đồng là");
    a.p.load("uniqueLocalId");
    await ctx.sync();
    const id0 = a.p.uniqueLocalId;
    const i0 = a.i;
    ctx.document.body.paragraphs.getFirst().insertParagraph("Đoạn chèn thử phía trên 1", "After");
    ctx.document.body.paragraphs.getFirst().insertParagraph("Đoạn chèn thử phía trên 2", "After");
    await ctx.sync();
    const b = await find(ctx, "Giá trị hợp đồng là");
    b.p.load("uniqueLocalId");
    const byId = ctx.document.getParagraphByUniqueLocalId(id0);
    byId.load("text");
    await ctx.sync();
    const h = byId.search("5.000.000.000", { matchCase: true });
    h.load("items");
    await ctx.sync();
    h.items[0].insertText("6.000.000.000", "Replace");
    await ctx.sync();
    const again = ctx.document.getParagraphByUniqueLocalId(id0);
    again.load("text,uniqueLocalId");
    await ctx.sync();
    return { chiSoTruoc: i0, chiSoSau: b.i, idGiuNguyen: b.p.uniqueLocalId === id0, timTheoId: byId.text.slice(0, 40), sauKhiSuaChu: again.text.slice(0, 50), idSauSua: again.uniqueLocalId === id0 };
  })
);

// (d) Chèn đoạn kế thừa đánh số: sau khoản 1.1 (gắn kiểu tiêu đề) và sau khoản 1.1 (numPr trên đoạn), đang TrackAll
const insertAfter = (label, start) =>
  test(label, () =>
    Word.run(async (ctx) => {
      const { p, i, ps } = await find(ctx, start);
      await setMode(ctx, "TrackAll");
      const np = p.insertParagraph("Khoản mới chèn thử.", "After");
      np.load("text,style,isListItem,uniqueLocalId");
      await ctx.sync();
      const all = ctx.document.body.paragraphs;
      all.load("items/text,items/style,items/listItemOrNullObject/listString,items/listItemOrNullObject/level");
      await ctx.sync();
      await setMode(ctx, "Off");
      const row = (k) => {
        const x = all.items[k];
        return `${x.listItemOrNullObject.isNullObject ? "∅" : x.listItemOrNullObject.listString} [${x.style}] ${x.text.slice(0, 34)}`;
      };
      return { neo: i, truocDo: ps.items.length, sauDo: all.items.length, quanhChoChen: [row(i - 1), row(i), row(i + 1), row(i + 2), row(i + 3)] };
    })
  );
await insertAfter("d1. chèn sau khoản gắn kiểu tiêu đề", "Bên A giao cho Bên B thực hiện");
await insertAfter("d2. chèn sau khoản đánh số trên đoạn", "Được kiểm tra tiến độ");

// (d3) Đổi cấp đoạn vừa chèn (numPr trên đoạn): level 1 → 2 rồi trả lại
await test("d3. đổi cấp bằng listItem.level", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const idx = ps.items.map((p, i) => (p.text === "Khoản mới chèn thử." ? i : -1)).filter((i) => i >= 0);
    const res = {};
    for (const i of idx) {
      const p = ps.items[i];
      const li = p.listItemOrNullObject;
      li.load("listString,level");
      p.load("style");
      await ctx.sync();
      const before = li.isNullObject ? "∅" : `${li.listString} lvl${li.level}`;
      let after = "";
      try {
        li.level = li.level + 1;
        await ctx.sync();
        const li2 = p.listItemOrNullObject;
        li2.load("listString,level");
        p.load("style");
        await ctx.sync();
        after = `${li2.listString} lvl${li2.level} [${p.style}]`;
      } catch (e) {
        after = "LỖI " + e.message;
      }
      res[`đoạn ${i} [${p.style}]`] = `${before} → ${after}`;
    }
    return res;
  })
);

// (e) Dữ liệu cho Hoàn tác: sửa đổi do add-in tạo có tác giả, ngày thế nào; từ chối riêng các sửa đổi đó
await test("e. hoàn tác bằng từ chối sửa đổi của add-in", () =>
  Word.run(async (ctx) => {
    const all = ctx.document.body.getTrackedChanges();
    all.load("items/type,items/author,items/text,items/date");
    await ctx.sync();
    const now = new Date().toISOString();
    const mine = all.items.filter((t) => t.author !== "Đối tác");
    const rows = tcList({ items: all.items });
    const n0 = all.items.length;
    mine.forEach((t) => t.reject());
    await ctx.sync();
    const left = ctx.document.body.getTrackedChanges();
    left.load("items/type,items/author,items/text");
    const { p } = await find(ctx, "Thời hạn thuê là");
    const cur = p.getText();
    const b = await find(ctx, "Bên B thanh toán trong");
    const cur2 = b.p.getText();
    await ctx.sync();
    return { gioMay: now, truoc: n0, tatCa: rows, cuaAddIn: mine.length, conLai: tcList2(left), doanDoiTacSauKhiTuChoi: cur.value, doanThanhToanSauKhiTuChoi: cur2.value };
    function tcList2(c) {
      return c.items.map((t) => `${t.type} ${JSON.stringify(t.text)} · ${who(t.author)}`);
    }
  })
);
return out;
