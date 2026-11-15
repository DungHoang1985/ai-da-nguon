// Thăm dò Đợt 1A (bước 1): so 4 đường đọc số điều khoản trên văn bản thử có 4 kiểu đánh số, và các lối đọc đoạn khi có sửa đổi được theo dõi.
// Văn bản thử: node kiem-thu/docx-thu.mjs tham-do <tệp> rồi mở bằng Word kèm cổng gỡ lỗi.
// Chạy: node kiem-thu/cdp.mjs eval 9229 kiem-thu/cdp-bai/tham-do-1-doc-so.js
const sup = (set, v) => Office.context.requirements.isSetSupported(set, v);
const out = { sets: {}, doan: [], duong: {}, giay: {} };
for (const v of ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]) out.sets["WordApi " + v] = sup("WordApi", v);
for (const v of ["1.1", "1.2", "1.3", "1.4", "1.5"]) out.sets["WordApiDesktop " + v] = sup("WordApiDesktop", v);
const timed = async (name, fn) => {
  const t0 = performance.now();
  try {
    out.duong[name] = await fn();
  } catch (e) {
    out.duong[name] = { LOI: (e && e.message) || String(e), o: e && e.debugInfo && e.debugInfo.errorLocation };
  }
  out.giay[name] = Math.round(performance.now() - t0) / 1000;
};

// Nền: mọi đoạn với thuộc tính cơ bản
let n = 0;
await timed("0. nền (text, style, outlineLevel…)", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text,items/style,items/styleBuiltIn,items/isListItem,items/outlineLevel,items/tableNestingLevel,items/uniqueLocalId");
    await ctx.sync();
    n = ps.items.length;
    out.doan = ps.items.map((p, i) => ({ i, text: p.text.slice(0, 60), style: p.style, builtIn: p.styleBuiltIn, isListItem: p.isListItem, outline: p.outlineLevel, tbl: p.tableNestingLevel }));
    return `${n} đoạn`;
  })
);

// Đường 1: Paragraph.listItemOrNullObject (WordApi 1.3)
await timed("1. listItemOrNullObject", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items");
    await ctx.sync();
    const li = ps.items.map((p) => {
      const x = p.listItemOrNullObject;
      x.load("listString,level,siblingIndex");
      return x;
    });
    await ctx.sync();
    li.forEach((x, i) => (out.doan[i].d1 = x.isNullObject ? null : `${JSON.stringify(x.listString)} lvl${x.level} sib${x.siblingIndex}`));
    return "xong";
  })
);

// Đường 1b: nạp gộp trong một lần sync
await timed("1b. nạp gộp items/listItemOrNullObject", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/isListItem,items/listItemOrNullObject/listString,items/listItemOrNullObject/level");
    await ctx.sync();
    return ps.items.map((p) => (p.listItemOrNullObject.isNullObject ? null : p.listItemOrNullObject.listString)).filter((x) => x !== null).length + " đoạn có số";
  })
);

// Đường 2: Range.listFormat (WordApiDesktop 1.3)
await timed("2. Range.listFormat", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items");
    await ctx.sync();
    const lf = ps.items.map((p) => {
      const x = p.getRange("Whole").listFormat;
      x.load("listString,listLevelNumber,listValue,listType");
      return x;
    });
    await ctx.sync();
    lf.forEach((x, i) => (out.doan[i].d2 = x.listType === "ListNoNumbering" && !x.listString ? null : `${JSON.stringify(x.listString)} lvl${x.listLevelNumber} val${x.listValue} ${x.listType}`));
    return "xong";
  })
);

// Đường 3: Document.listParagraphs (WordApiDesktop 1.4)
await timed("3. Document.listParagraphs", () =>
  Word.run(async (ctx) => {
    const lp = ctx.document.listParagraphs;
    lp.load("items/text,items/uniqueLocalId");
    await ctx.sync();
    const li = lp.items.map((p) => {
      const x = p.listItemOrNullObject;
      x.load("listString,level");
      return x;
    });
    await ctx.sync();
    return { soDoan: lp.items.length, muc: lp.items.map((p, i) => `${li[i].isNullObject ? "∅" : li[i].listString} | ${p.text.slice(0, 30)}`) };
  })
);

// Đường 4: Document.getCrossReferenceItems (WordApiDesktop 1.4)
for (const kind of ["NumberedItem", "Heading"]) {
  await timed(`4. getCrossReferenceItems(${kind})`, () =>
    Word.run(async (ctx) => {
      const r = ctx.document.getCrossReferenceItems(kind);
      await ctx.sync();
      return r.value;
    })
  );
}

// Trường trong văn bản
await timed("trường (Body.fields)", () =>
  Word.run(async (ctx) => {
    const fs = ctx.document.body.fields;
    fs.load("items/code,items/type");
    await ctx.sync();
    const rs = fs.items.map((f) => {
      const r = f.result;
      r.load("text");
      return r;
    });
    await ctx.sync();
    return fs.items.map((f, i) => `${f.type}: ${f.code.trim()} → ${JSON.stringify(rs[i].text)}`);
  })
);

// Đọc đoạn có sửa đổi: .text so với getText / getReviewedText
await timed("đọc khi có sửa đổi", () =>
  Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    const from = ps.items.findIndex((p) => p.text.startsWith("[E]"));
    const rows = ps.items.slice(from + 1, from + 5).map((p, k) => {
      const tc = p.getTrackedChanges();
      tc.load("items/type,items/author,items/text,items/date");
      return {
        i: from + 1 + k,
        text: p.text,
        getText: p.getText(),
        getTextDel: p.getText({ IncludeTextMarkedAsDeleted: true, IncludeHiddenText: false }),
        cur: p.getReviewedText("Current"),
        orig: p.getReviewedText("Original"),
        tc,
      };
    });
    await ctx.sync();
    ctx.document.load("changeTrackingMode");
    const all = ctx.document.body.getTrackedChanges();
    all.load("items");
    await ctx.sync();
    return {
      mode: ctx.document.changeTrackingMode,
      tongSuaDoi: all.items.length,
      doan: rows.map((r) => ({
        i: r.i,
        ".text": r.text,
        "getText()": r.getText.value,
        "getText(+đã xóa)": r.getTextDel.value,
        "getReviewedText(Current)": r.cur.value,
        "getReviewedText(Original)": r.orig.value,
        suaDoi: r.tc.items.map((t) => `${t.type} "${t.text}" bởi ${t.author} lúc ${new Date(t.date).toISOString()}`),
      })),
    };
  })
);
return out;
