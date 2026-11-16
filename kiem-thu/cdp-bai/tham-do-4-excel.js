// Đợt 4: thăm dò các API Excel trước khi viết công cụ (biểu đồ, Pivot, thiết lập in, tên vùng, R1C1, làm sạch, nhóm, bảo vệ, giải ngược).
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-bai/tham-do-4-excel.js
// Chỉ đụng hai sheet thử AIDN_TD4 và AIDN_TD4_P, xong thì xóa.
const out = {};
const S = "AIDN_TD4";
const P = "AIDN_TD4_P";
const t = async (name, fn) => {
  const t0 = performance.now();
  try {
    out[name] = { ok: true, kq: await fn(), ms: Math.round(performance.now() - t0) };
  } catch (e) {
    out[name] = { ok: false, loi: `${e?.code || ""} ${e?.message || e}`.trim(), cho: e?.debugInfo?.errorLocation };
  }
};
const sets = {};
for (const v of ["1.9", "1.10", "1.11", "1.12", "1.13", "1.14", "1.15", "1.16", "1.17", "1.18", "1.19", "1.20"]) sets[v] = Office.context.requirements.isSetSupported("ExcelApi", v);
out.boApi = sets;

await Excel.run(async (ctx) => {
  for (const n of [S, P]) {
    const w = ctx.workbook.worksheets.getItemOrNullObject(n);
    await ctx.sync();
    if (!w.isNullObject) w.delete();
  }
  const ws = ctx.workbook.worksheets.add(S);
  ws.getRange("A1:D5").values = [
    ["Nhóm", "Quý", "Doanh thu", "Chi phí"],
    ["A", "Q1", 100, 60],
    ["B", "Q1", 200, 120],
    ["A", "Q2", 150, 80],
    ["B", "Q2", 250, 140],
  ];
  ws.activate();
  await ctx.sync();
});

await t("bieuDo_doiKieu_tieuDe_nguon", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const ch = ws.charts.add("ColumnClustered", ws.getRange("A1:C3"), "Auto");
    ch.title.text = "Cũ";
    ch.load("name,chartType");
    await ctx.sync();
    const truoc = ch.chartType;
    ch.chartType = "Line";
    ch.title.text = "Mới";
    ch.setData(ws.getRange("C1:D5"), "Columns");
    ch.axes.categoryAxis.title.text = "Trục X";
    ch.axes.valueAxis.title.text = "Trục Y";
    ch.legend.position = "Bottom";
    ch.load("chartType");
    ch.title.load("text");
    ch.series.load("items/name");
    await ctx.sync();
    const nguon = [];
    for (const s of ch.series.items) {
      try {
        const v = s.getDimensionDataSourceString("Values");
        const c = s.getDimensionDataSourceString("Categories");
        await ctx.sync();
        nguon.push({ ten: s.name, values: v.value, categories: c.value });
      } catch (e) {
        nguon.push({ ten: s.name, loi: e.code || e.message });
      }
    }
    return { ten: ch.name, truoc, sau: ch.chartType, tieuDe: ch.title.text, chuoi: nguon };
  })
);

await t("bieuDo_lietKe", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ws.charts.load("items/name,items/chartType,items/id");
    await ctx.sync();
    const r = [];
    for (const c of ws.charts.items) {
      c.title.load("text,visible");
      await ctx.sync();
      r.push({ ten: c.name, kieu: c.chartType, tieuDe: c.title.text });
    }
    return r;
  })
);

let pivotName = "PivotTD4";
await t("pivot_tao", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const dw = ctx.workbook.worksheets.add(P);
    const pt = dw.pivotTables.add(pivotName, ws.getRange("A1:D5"), dw.getRange("A3"));
    pt.rowHierarchies.add(pt.hierarchies.getItem("Nhóm"));
    pt.dataHierarchies.add(pt.hierarchies.getItem("Doanh thu"));
    await ctx.sync();
    return "xong";
  })
);
await t("pivot_them_bot_truong_ham_loc_lamMoi", () =>
  Excel.run(async (ctx) => {
    const pt = ctx.workbook.pivotTables.getItem(pivotName);
    pt.columnHierarchies.add(pt.hierarchies.getItem("Quý"));
    const dh = pt.dataHierarchies.add(pt.hierarchies.getItem("Chi phí"));
    dh.summarizeBy = "Average";
    await ctx.sync();
    pt.rowHierarchies.load("items/name");
    pt.columnHierarchies.load("items/name");
    pt.dataHierarchies.load("items/name,items/summarizeBy");
    pt.filterHierarchies.load("items/name");
    pt.hierarchies.load("items/name");
    await ctx.sync();
    const sauThem = {
      hang: pt.rowHierarchies.items.map((x) => x.name),
      cot: pt.columnHierarchies.items.map((x) => x.name),
      giaTri: pt.dataHierarchies.items.map((x) => `${x.name}|${x.summarizeBy}`),
      tatCa: pt.hierarchies.items.map((x) => x.name),
    };
    // lọc theo mục: chỉ giữ nhóm A
    let loc = "";
    try {
      const f = pt.rowHierarchies.getItem("Nhóm").fields.getItem("Nhóm");
      f.applyFilter({ manualFilter: { selectedItems: ["A"] } });
      await ctx.sync();
      const rng = pt.layout.getRange();
      rng.load("values,address");
      await ctx.sync();
      loc = JSON.stringify(rng.values).slice(0, 300);
      f.clearAllFilters();
      await ctx.sync();
    } catch (e) {
      loc = "LỖI " + (e.code || e.message);
    }
    // bớt trường cột, đổi hàm
    pt.columnHierarchies.remove(pt.columnHierarchies.getItem("Quý"));
    pt.dataHierarchies.remove(pt.dataHierarchies.items[1]);
    await ctx.sync();
    // đổi dữ liệu nguồn rồi làm mới
    ctx.workbook.worksheets.getItem(S).getRange("C2").values = [[1000]];
    await ctx.sync();
    const r1 = pt.layout.getRange();
    r1.load("values");
    await ctx.sync();
    const truocLamMoi = JSON.stringify(r1.values);
    pt.refresh();
    await ctx.sync();
    const r2 = pt.layout.getRange();
    r2.load("values,address");
    pt.columnHierarchies.load("items/name");
    pt.dataHierarchies.load("items/name");
    await ctx.sync();
    return { sauThem, loc, truocLamMoi, sauLamMoi: JSON.stringify(r2.values), vung: r2.address, cotCon: pt.columnHierarchies.items.length, giaTriCon: pt.dataHierarchies.items.map((x) => x.name) };
  })
);
await t("pivot_lietKe_caSo", () =>
  Excel.run(async (ctx) => {
    const pts = ctx.workbook.pivotTables;
    pts.load("items/name");
    await ctx.sync();
    const r = [];
    for (const p of pts.items) {
      p.worksheet.load("name");
      const src = p.getDataSourceString ? p.getDataSourceString() : null;
      await ctx.sync();
      r.push({ ten: p.name, sheet: p.worksheet.name, nguon: src ? src.value : "(không có getDataSourceString)" });
    }
    return r;
  })
);

await t("in_dat_va_doc_lai", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const pl = ws.pageLayout;
    pl.load("orientation,paperSize,zoom,centerHorizontally,printGridlines,topMargin,bottomMargin,leftMargin,rightMargin");
    await ctx.sync();
    const truoc = { huong: pl.orientation, kho: pl.paperSize, zoom: pl.zoom, tren: pl.topMargin, trai: pl.leftMargin };
    pl.orientation = "Landscape";
    pl.paperSize = "A4";
    pl.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
    pl.setPrintArea(ws.getRange("A1:D5"));
    pl.setPrintTitleRows(ws.getRange("1:1"));
    pl.setPrintMargins("Centimeters", { top: 1.5, bottom: 1.5, left: 1, right: 1 });
    pl.centerHorizontally = true;
    pl.headersFooters.defaultForAllPages.centerHeader = "Bảng tiến độ";
    pl.headersFooters.defaultForAllPages.centerFooter = "Trang &P / &N";
    await ctx.sync();
    pl.load("orientation,paperSize,zoom,centerHorizontally,topMargin,leftMargin");
    const pa = pl.getPrintAreaOrNullObject();
    pa.load("address");
    const tr = pl.getPrintTitleRowsOrNullObject();
    tr.load("address");
    pl.headersFooters.defaultForAllPages.load("centerHeader,centerFooter");
    await ctx.sync();
    return {
      truoc,
      sau: { huong: pl.orientation, kho: pl.paperSize, zoom: pl.zoom, tren: pl.topMargin, trai: pl.leftMargin, giua: pl.centerHorizontally },
      vungIn: pa.isNullObject ? null : pa.address,
      tieuDeLap: tr.isNullObject ? null : tr.address,
      dau: pl.headersFooters.defaultForAllPages.centerHeader,
      chan: pl.headersFooters.defaultForAllPages.centerFooter,
    };
  })
);
await t("in_zoom_ve_null_va_bo_vung_in", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const pl = ws.pageLayout;
    pl.zoom = { scale: 100 };
    const pa0 = pl.getPrintAreaOrNullObject();
    await ctx.sync();
    let bo = "";
    try {
      pl.setPrintArea(ws.getRange("A1:B2"));
      await ctx.sync();
      // thử bỏ vùng in: không có hàm clear → thử chuỗi rỗng
      pl.setPrintArea("");
      await ctx.sync();
      bo = "setPrintArea('') được";
    } catch (e) {
      bo = "setPrintArea('') lỗi " + (e.code || e.message);
    }
    pl.load("zoom");
    const pa = pl.getPrintAreaOrNullObject();
    pa.load("address");
    await ctx.sync();
    return { zoom: pl.zoom, bo, vungInCon: pa.isNullObject ? null : pa.address };
  })
);

await t("tenVung", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ctx.workbook.names.add("Lãi_suất_TD4", ws.getRange("C2"), "giả định thử");
    await ctx.sync();
    ws.getRange("F1").formulas = [["=Lãi_suất_TD4*2"]];
    const f = ws.getRange("F1");
    f.load("values");
    ctx.workbook.names.load("items/name,items/formula,items/value,items/visible,items/comment,items/scope");
    await ctx.sync();
    const ds = ctx.workbook.names.items.map((n) => ({ ten: n.name, ct: n.formula, gt: n.value, hien: n.visible }));
    const trung = ctx.workbook.names.getItemOrNullObject("Lãi_suất_TD4");
    await ctx.sync();
    ctx.workbook.names.getItem("Lãi_suất_TD4").delete();
    await ctx.sync();
    const f2 = ws.getRange("F1");
    f2.load("values");
    await ctx.sync();
    ws.getRange("F1").clear();
    await ctx.sync();
    return { F1: f.values[0][0], ds, coSan: !trung.isNullObject, F1sauXoaTen: f2.values[0][0] };
  })
);

await t("r1c1_va_hangSo", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ws.getRange("H1:K4").formulas = [
      ["Năm 1", "Năm 2", "Năm 3", "Năm 4"],
      [100, "=H2*1.1", "=I2*1.1", "=J2*1.1"],
      [50, "=H3*1.05", 57, "=J3*1.05"],
      ["=H2-H3", "=I2-I3", "=J2-J4", "=K2-K3"],
    ];
    const r = ws.getRange("H1:K4");
    r.load("formulasR1C1,formulas,values,valueTypes");
    await ctx.sync();
    return { r1c1: r.formulasR1C1, kieu: r.valueTypes };
  })
);

await t("giaiNguoc_ghi_roi_doc_cung_sync", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ctx.workbook.application.load("calculationMode");
    await ctx.sync();
    const mode = ctx.workbook.application.calculationMode;
    const ket = [];
    const t0 = performance.now();
    for (const x of [100, 120, 140]) {
      ws.getRange("H2").values = [[x]];
      const o = ws.getRange("K4");
      o.load("values");
      await ctx.sync();
      ket.push([x, o.values[0][0]]);
    }
    ws.getRange("H2").values = [[100]];
    await ctx.sync();
    return { cheDoTinh: mode, ket, msMoiVong: Math.round((performance.now() - t0) / 3) };
  })
);
await t("giaiNguoc_khi_tinh_tay", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const app = ctx.workbook.application;
    app.calculationMode = "Manual";
    await ctx.sync();
    ws.getRange("H2").values = [[200]];
    const o = ws.getRange("K4");
    o.load("values");
    await ctx.sync();
    const khongTinh = o.values[0][0];
    ws.calculate(true);
    const o2 = ws.getRange("K4");
    o2.load("values");
    await ctx.sync();
    const sauCalculate = o2.values[0][0];
    ws.getRange("H2").values = [[100]];
    app.calculationMode = "Automatic";
    await ctx.sync();
    return { khongTinh, sauCalculate };
  })
);

await t("lamSach_xoaTrung_thayHangLoat_tuDien", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ws.getRange("M1:N6").values = [
      ["Mã", "Tình trạng"],
      ["HS01", "Chờ"],
      ["HS02", "Chờ"],
      ["HS01", "Chờ"],
      ["HS03", "Xong"],
      ["HS02", "Chờ"],
    ];
    await ctx.sync();
    const rd = ws.getRange("M1:N6").removeDuplicates([0], true);
    rd.load("removed,uniqueRemaining");
    await ctx.sync();
    const n = ws.getRange("M1:N6").replaceAll("Chờ", "Đang xử lý", { completeMatch: false, matchCase: false });
    await ctx.sync();
    ws.getRange("P1:P2").values = [["01/12/2026"], ["02/12/2026"]];
    ws.getRange("Q1").values = [["01/12/2026"]];
    await ctx.sync();
    ws.getRange("P1:P2").autoFill(ws.getRange("P1:P6"), "FillDefault");
    let fillDays = "";
    try {
      ws.getRange("Q1").autoFill(ws.getRange("Q1:Q6"), "FillWeekdays");
      await ctx.sync();
      fillDays = "FillWeekdays được";
    } catch (e) {
      fillDays = "FillWeekdays lỗi " + (e.code || e.message);
    }
    const a = ws.getRange("M1:Q6");
    a.load("values,text");
    await ctx.sync();
    return { xoa: rd.removed, con: rd.uniqueRemaining, thay: n.value, text: a.text, fillDays };
  })
);

await t("nhom_hang_cot", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ws.getRange("2:3").group("ByRows");
    ws.getRange("C:D").group("ByColumns");
    await ctx.sync();
    ws.getRange("2:3").hideGroupDetails("ByRows");
    await ctx.sync();
    const r = ws.getRange("A2");
    r.load("rowHidden");
    await ctx.sync();
    const an = r.rowHidden;
    ws.getRange("2:3").showGroupDetails("ByRows");
    ws.getRange("2:3").ungroup("ByRows");
    ws.getRange("C:D").ungroup("ByColumns");
    await ctx.sync();
    return { anKhiThuGon: an };
  })
);

await t("baoVe_khoa_o_cong_thuc", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const used = ws.getUsedRange();
    used.format.protection.load("locked");
    ws.protection.load("protected");
    await ctx.sync();
    const truoc = { khoa: used.format.protection.locked, baoVe: ws.protection.protected };
    used.format.protection.locked = false;
    const fs = used.getSpecialCellsOrNullObject("Formulas");
    fs.load("address,cellCount");
    await ctx.sync();
    if (!fs.isNullObject) fs.format.protection.locked = true;
    ws.protection.protect({ allowFormatCells: true, allowFormatColumns: true, allowFormatRows: true, allowAutoFilter: true, allowSort: true, selectionMode: "Normal" });
    await ctx.sync();
    let ghiCongThuc = "";
    try {
      ws.getRange("I2").values = [[1]];
      await ctx.sync();
      ghiCongThuc = "GHI ĐƯỢC (sai)";
    } catch (e) {
      ghiCongThuc = "bị chặn: " + (e.code || e.message);
    }
    let ghiNhap = "";
    try {
      await Excel.run(async (c2) => {
        c2.workbook.worksheets.getItem(S).getRange("H2").values = [[101]];
        await c2.sync();
        c2.workbook.worksheets.getItem(S).getRange("H2").values = [[100]];
        await c2.sync();
      });
      ghiNhap = "ghi được";
    } catch (e) {
      ghiNhap = "LỖI " + (e.code || e.message);
    }
    return { truoc, oCongThuc: fs.isNullObject ? 0 : fs.cellCount, diaChi: fs.isNullObject ? "" : fs.address.slice(0, 120), ghiCongThuc, ghiNhap };
  })
);
await t("baoVe_bo", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    ws.protection.unprotect();
    await ctx.sync();
    ws.protection.load("protected");
    const u = ws.getUsedRange();
    u.format.protection.load("locked");
    await ctx.sync();
    return { baoVe: ws.protection.protected, khoaHonHop: u.format.protection.locked };
  })
);

await t("lienKetNgoai_api", () =>
  Excel.run(async (ctx) => {
    if (!ctx.workbook.linkedWorkbooks) return "không có workbook.linkedWorkbooks";
    ctx.workbook.linkedWorkbooks.load("items/id");
    await ctx.sync();
    return { so: ctx.workbook.linkedWorkbooks.items.length };
  })
);

await t("toMau_va_doc_lai_fill", () =>
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(S);
    const r = ws.getRange("A2:B2");
    r.format.fill.load("color");
    await ctx.sync();
    const truoc = r.format.fill.color;
    r.format.fill.color = "#FFF2CC";
    await ctx.sync();
    r.format.fill.clear();
    await ctx.sync();
    r.format.fill.load("color");
    await ctx.sync();
    return { truoc, sauClear: r.format.fill.color };
  })
);

await Excel.run(async (ctx) => {
  for (const n of [P, S]) {
    const w = ctx.workbook.worksheets.getItemOrNullObject(n);
    await ctx.sync();
    if (!w.isNullObject) {
      try {
        w.protection.unprotect();
      } catch {}
      w.delete();
    }
  }
  await ctx.sync();
}).catch((e) => (out.donDep = "LỖI " + e.message));
return out;
