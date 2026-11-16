// Đợt 4 (Chạy thật): giao một việc cho AI trong Excel rồi ĐỌC LẠI sổ tính bằng mã.
// Tham số ở tệp ghép phía trước: window.__aidnBai = { yeuCau, nguon?, model?, traLoi?: "allow" | "all" | "deny" | [mẫu, trả lời][], doc?: ["Sheet!A1:B2", …], hoanTac?: true }
// Trả: nguồn và model lúc bắt đầu (để trả lại), các lần hỏi, số lần gọi AI, phần cuối khung, và trạng thái sổ tính đọc lại bằng mã.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-lib.js <tệp tham số> kiem-thu/cdp-bai/dot-4-excel-bai.js
const bai = window.__aidnBai || {};
delete window.__aidnBai;
if (!bai.yeuCau) return { loi: "Thiếu window.__aidnBai.yeuCau" };

const docSo = () =>
  Excel.run(async (ctx) => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    ctx.workbook.names.load("items/name,items/formula");
    await ctx.sync();
    const out = { sheets: [], tenVung: ctx.workbook.names.items.map((n) => `${n.name} ${n.formula}`) };
    for (const ws of wss.items) {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load("address");
      ws.charts.load("items/name,items/chartType");
      ws.pivotTables.load("items/name");
      ws.freezePanes.getLocationOrNullObject().load("address");
      const pl = ws.pageLayout;
      pl.load("orientation,paperSize,zoom");
      const pa = pl.getPrintAreaOrNullObject();
      pa.load("address");
      const tr = pl.getPrintTitleRowsOrNullObject();
      tr.load("address");
      ws.protection.load("protected");
      await ctx.sync();
      const charts = [];
      for (const c of ws.charts.items) {
        c.title.load("text");
        await ctx.sync();
        charts.push(`${c.name}|${c.chartType}|${c.title.text}`);
      }
      out.sheets.push({
        ten: ws.name,
        vung: u.isNullObject ? null : u.address,
        bieuDo: charts,
        pivot: ws.pivotTables.items.map((p) => p.name),
        in: { huong: pl.orientation, kho: pl.paperSize, vuaNgang: pl.zoom?.horizontalFitToPages ?? null, vuaDoc: pl.zoom?.verticalFitToPages ?? null, vungIn: pa.isNullObject ? null : pa.address, tieuDeLap: tr.isNullObject ? null : tr.address },
        baoVe: ws.protection.protected,
      });
    }
    const doc = {};
    for (const addr of bai.doc || []) {
      try {
        const [sh, a] = addr.split("!");
        const r = ctx.workbook.worksheets.getItem(sh).getRange(a);
        r.load("values,formulas");
        r.format.fill.load("color");
        r.format.font.load("bold");
        await ctx.sync();
        doc[addr] = { values: r.values, formulas: r.formulas.map((row) => row.map((f) => (typeof f === "string" && f.startsWith("=") ? f : null))), nen: r.format.fill.color, dam: r.format.font.bold };
      } catch (e) {
        doc[addr] = "LỖI " + (e.code || e.message);
      }
    }
    return { ...out, doc };
  });

await toChat();
btn(/^Hội thoại mới/)?.click();
await sleep(500);
const truoc = label(btn(/^Model đang dùng/) || document.body).slice(0, 120);
let chon = null;
if (bai.nguon && bai.model) chon = await pick(bai.nguon, new RegExp(bai.model, "i"));
if (chon?.loi) return { loi: chon.loi, coSan: chon.coSan, modelTruoc: truoc };
const approve = (q) => {
  if (Array.isArray(bai.traLoi)) {
    for (const [re, ans] of bai.traLoi) if (new RegExp(re, "i").test(q)) return ans;
    return "allow";
  }
  return bai.traLoi || "allow";
};
const r = await send(bai.yeuCau, { approve, timeout: bai.timeout || 240000, tail: 2600 });
const khung = r.khung || "";
const goi = /(\d+) lần gọi AI/.exec(khung);
let hoanTac = null;
let soTruocHoanTac = null;
if (bai.hoanTac) {
  soTruocHoanTac = await docSo();
  const u = btn(/^Hoàn tác/);
  if (u) {
    u.click();
    await sleep(2500);
    hoanTac = "đã bấm: " + label(u).slice(0, 60);
  } else hoanTac = "không thấy nút Hoàn tác";
}
return {
  modelTruoc: truoc,
  modelThu: chon?.daChon || truoc,
  phienBan: (document.body.innerText.match(/v\d+\.\d+(\.\d+)?/) || [""])[0],
  giay: r.giay,
  daHoi: r.daHoi,
  lanGoiAI: goi ? Number(goi[1]) : null,
  hoanTac,
  soTruocHoanTac,
  khung: khung.replace(/\n{2,}/g, "\n"),
  so: await docSo(),
};
