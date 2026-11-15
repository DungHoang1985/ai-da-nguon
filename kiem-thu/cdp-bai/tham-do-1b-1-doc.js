// Thăm dò Đợt 1B (bước 2b, 2c): API nào đọc được chú thích, đầu/chân trang, bookmark, trường, ô điền mẫu, hộp văn bản,
// thể thức trang, phông chữ trên máy này. Văn bản thử: node kiem-thu/soat-thu.mjs tham-do <tệp tham-do-1b.docx>.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-1b-1-doc.js
const sup = (set, v) => Office.context.requirements.isSetSupported(set, v);
const out = { sets: {}, kq: {}, giay: {} };
for (const v of ["1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]) out.sets["WordApi " + v] = sup("WordApi", v);
for (const v of ["1.1", "1.2", "1.3", "1.4", "1.5"]) out.sets["WordApiDesktop " + v] = sup("WordApiDesktop", v);
for (const v of ["1.3", "1.4", "1.5"]) out.sets["WordApiHiddenDocument " + v] = sup("WordApiHiddenDocument", v);
const cut = (s, n = 90) => (String(s ?? "").length > n ? String(s).slice(0, n) + "…" : String(s ?? ""));
const timed = async (name, fn) => {
  const t0 = performance.now();
  try {
    out.kq[name] = await fn();
  } catch (e) {
    out.kq[name] = { LOI: (e && e.message) || String(e), ma: e && e.code, o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
  out.giay[name] = Math.round(performance.now() - t0) / 1000;
};

// 0. Chữ của các đoạn có mốc chú thích, trường, ô điền mẫu, hộp văn bản: paragraph.text chứa gì
await timed("0. paragraph.text của đoạn có mốc", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const pick = (re) => ps.items.map((p, i) => [i, p.text]).filter(([, t]) => re.test(t)).map(([i, t]) => `[${i}] ${JSON.stringify(cut(t, 200))}`);
    return {
      soDoan: ps.items.length,
      tieuDe_hopVanBan: pick(/^HỢP ĐỒNG THUÊ/),
      chuThich_1_4: pick(/^“Tiền Thuê” là/),
      truongREF_10_2: pick(/đúng hạn theo Điều/),
      oDienMau_ky: pick(/^Ông /),
      oTrong: pick(/Ô điền mẫu còn trống/),
      chuThichCuoi: pick(/nhà thầu của mình trong/),
    };
  })
);

// 1. Chú thích cuối trang, cuối văn bản (WordApi 1.5)
await timed("1. body.footnotes / endnotes", () =>
  Word.run(async (ctx) => {
    const fn = ctx.document.body.footnotes;
    const en = ctx.document.body.endnotes;
    fn.load("items/type,items/body/text");
    en.load("items/type,items/body/text");
    await ctx.sync();
    const refs = [...fn.items, ...en.items].map((n) => {
      const p = n.reference.paragraphs.getFirst();
      p.load("text,uniqueLocalId");
      return p;
    });
    await ctx.sync();
    const one = (n, k) => ({ loai: n.type, chu: cut(n.body.text, 120), doanNeo: cut(refs[k].text, 50), id: !!refs[k].uniqueLocalId });
    return { footnotes: fn.items.map(one), endnotes: en.items.map((n, k) => one(n, fn.items.length + k)) };
  })
);

// 2. Đầu trang, chân trang theo từng phần (WordApi 1.1)
await timed("2. sections getHeader/getFooter", () =>
  Word.run(async (ctx) => {
    const secs = ctx.document.sections;
    secs.load("items");
    await ctx.sync();
    const res = [];
    for (const [k, s] of secs.items.entries()) {
      const row = { phan: k };
      for (const t of ["Primary", "FirstPage", "EvenPages"]) {
        const h = s.getHeader(t);
        const f = s.getFooter(t);
        h.load("text");
        f.load("text");
        try {
          await ctx.sync();
          row[t] = { dau: cut(h.text, 120), chan: cut(f.text, 120) };
        } catch (e) {
          row[t] = { LOI: e.message };
        }
      }
      res.push(row);
    }
    return res;
  })
);
await timed("2b. chân trang: đoạn và trường", () =>
  Word.run(async (ctx) => {
    const f = ctx.document.sections.getFirst().getFooter("Primary");
    f.paragraphs.load("items/text");
    f.fields.load("items/code,items/type");
    await ctx.sync();
    return { doan: f.paragraphs.items.map((p) => cut(p.text, 120)), truong: f.fields.items.map((x) => `${x.type}: ${x.code}`) };
  })
);

// 3. Bookmark
await timed("3a. Range.getBookmarks(includeHidden) (WordApi 1.4)", () =>
  Word.run(async (ctx) => {
    const r = ctx.document.body.getRange("Whole");
    const visible = r.getBookmarks(false, false);
    const all = r.getBookmarks(true, false);
    await ctx.sync();
    return { hien: visible.value, caAn: all.value };
  })
);
await timed("3b. document.getBookmarkRangeOrNullObject (WordApi 1.4)", () =>
  Word.run(async (ctx) => {
    const res = {};
    for (const name of ["GiaThue", "_Ref7001", "KhongCo"]) {
      const r = ctx.document.getBookmarkRangeOrNullObject(name);
      r.load("text");
      const p = r.paragraphs.getFirstOrNullObject();
      p.load("text");
      await ctx.sync();
      res[name] = r.isNullObject ? null : { chu: cut(r.text, 100), doan: p.isNullObject ? null : cut(p.text, 50) };
    }
    return res;
  })
);
await timed("3c. document.bookmarks (WordApiDesktop 1.4)", () =>
  Word.run(async (ctx) => {
    const b = ctx.document.bookmarks;
    b.load("items/name,items/start,items/end,items/storyType,items/isEmpty");
    await ctx.sync();
    return b.items.map((x) => `${x.name} ${x.start}-${x.end} ${x.storyType}${x.isEmpty ? " rỗng" : ""}`);
  })
);

// 4. Trường (WordApi 1.4)
await timed("4. body.fields", () =>
  Word.run(async (ctx) => {
    const f = ctx.document.body.fields;
    f.load("items/code,items/type,items/kind,items/result/text");
    await ctx.sync();
    const ps = f.items.map((x) => {
      const p = x.result.paragraphs.getFirstOrNullObject();
      p.load("text");
      return p;
    });
    await ctx.sync();
    return f.items.map((x, k) => ({ loai: x.type, ma: x.code.trim(), ketQua: x.result.text, doan: ps[k].isNullObject ? null : cut(ps[k].text, 40) }));
  })
);

// 5. Ô điền mẫu (WordApi 1.1; placeholderText 1.3?)
await timed("5. contentControls", () =>
  Word.run(async (ctx) => {
    const cc = ctx.document.contentControls;
    cc.load("items/id,items/title,items/tag,items/text,items/type,items/placeholderText,items/cannotEdit,items/appearance");
    await ctx.sync();
    const ps = cc.items.map((x) => {
      const p = x.paragraphs.getFirstOrNullObject();
      p.load("text");
      return p;
    });
    await ctx.sync();
    return cc.items.map((x, k) => ({ id: x.id, tieuDe: x.title, tag: x.tag, loai: x.type, chu: x.text, goiY: x.placeholderText, doan: ps[k].isNullObject ? null : cut(ps[k].text, 40) }));
  })
);
await timed("5b. getContentControls mọi loại (WordApi 1.5)", () =>
  Word.run(async (ctx) => {
    const cc = ctx.document.getContentControls({ types: ["RichText", "PlainText", "CheckBox", "DropDownList", "ComboBox", "DatePicker"] });
    cc.load("items/title,items/type,items/text");
    await ctx.sync();
    return cc.items.map((x) => `${x.type} “${x.title}” = ${JSON.stringify(x.text)}`);
  })
);

// 6. Hộp văn bản (WordApiDesktop 1.2)
await timed("6. body.shapes", () =>
  Word.run(async (ctx) => {
    const sh = ctx.document.body.shapes;
    sh.load("items/id,items/name,items/type");
    await ctx.sync();
    const bodies = sh.items.map((s) => {
      s.body.load("text");
      const anchor = s.body; // thân hộp
      return anchor;
    });
    await ctx.sync();
    return sh.items.map((s, k) => ({ id: s.id, ten: s.name, loai: s.type, chu: cut(bodies[k].text, 120) }));
  })
);
await timed("6b. hộp văn bản: đoạn neo", () =>
  Word.run(async (ctx) => {
    const sh = ctx.document.body.shapes;
    sh.load("items/type");
    await ctx.sync();
    const tb = sh.items.find((s) => s.type === "TextBox");
    if (!tb) return "không có TextBox";
    const ps = tb.body.paragraphs;
    ps.load("items/text");
    const par = tb.body.parentBodyOrNullObject;
    par.load("type");
    await ctx.sync();
    return { doanTrongHop: ps.items.map((p) => p.text), thanCha: par.isNullObject ? null : par.type };
  })
);

// 7. Thể thức trang (WordApiDesktop 1.3)
const PS = "topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight,paperSize,orientation,headerDistance,footerDistance,gutter,mirrorMargins,differentFirstPageHeaderFooter,oddAndEvenPagesHeaderFooter";
await timed("7a. section.pageSetup", () =>
  Word.run(async (ctx) => {
    const secs = ctx.document.sections;
    secs.load("items");
    await ctx.sync();
    const ps = secs.items.map((s) => {
      s.pageSetup.load(PS);
      return s.pageSetup;
    });
    await ctx.sync();
    return ps.map((p) => p.toJSON());
  })
);
await timed("7b. document.pageSetup", () =>
  Word.run(async (ctx) => {
    const p = ctx.document.pageSetup;
    p.load(PS);
    await ctx.sync();
    return p.toJSON();
  })
);

// 8. Phông và cỡ chữ
await timed("8a. paragraphs items/font (name, size)", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text,items/font/name,items/font/size,items/alignment,items/firstLineIndent,items/lineSpacing");
    await ctx.sync();
    const count = {};
    ps.items.forEach((p) => {
      const k = `${JSON.stringify(p.font.name)} ${p.font.size}`;
      count[k] = (count[k] || 0) + p.text.length;
    });
    return { theoKyTu: count, vaiDoan: ps.items.slice(0, 5).map((p) => `${JSON.stringify(p.font.name)} ${p.font.size} ${p.alignment} | ${cut(p.text, 40)}`) };
  })
);
await timed("8b. body.font", () =>
  Word.run(async (ctx) => {
    const f = ctx.document.body.font;
    f.load("name,size");
    await ctx.sync();
    return `${JSON.stringify(f.name)} ${f.size}`;
  })
);
return out;
