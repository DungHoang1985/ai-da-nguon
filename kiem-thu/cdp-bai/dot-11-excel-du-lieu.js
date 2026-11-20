// Đợt 11: dựng lại sổ tính thử về đúng trạng thái đầu trước mỗi bài (bài trước / sau so được với nhau mà không phải mở lại Excel).
// Xóa mọi sheet (kèm biểu đồ, PivotTable trên đó) và mọi tên vùng, rồi tạo sheet "Bán hàng": 24 dòng Tháng · Khu vực · Dự án · Số căn · Doanh thu.
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-lib.js kiem-thu/cdp-bai/dot-11-excel-du-lieu.js
return Excel.run(async (ctx) => {
  const wb = ctx.workbook;
  wb.worksheets.load("items/name");
  wb.names.load("items/name");
  await ctx.sync();
  const tam = wb.worksheets.add("AIDN_tam_" + Date.now());
  await ctx.sync();
  for (const w of wb.worksheets.items) w.delete();
  for (const n of wb.names.items) n.delete();
  await ctx.sync();
  const ws = wb.worksheets.add("Bán hàng");
  const kv = [
    ["Hà Nội", "Ánh Dương"],
    ["Hải Phòng", "Cảng Xanh"],
    ["Đà Nẵng", "Sông Hàn"],
    ["TP.HCM", "Ánh Dương"],
  ];
  const rows = [["Tháng", "Khu vực", "Dự án", "Số căn", "Doanh thu (tỷ đồng)"]];
  for (let m = 1; m <= 6; m++)
    kv.forEach(([k, d], i) => {
      const can = 4 + ((m * 7 + i * 5) % 9);
      rows.push([`T${m}/2026`, k, d, can, Math.round(can * (2.1 + i * 0.6 + m * 0.05) * 10) / 10]);
    });
  ws.getRange(`A1:E${rows.length}`).values = rows;
  ws.getRange("A1:E1").format.font.bold = true;
  ws.getRange("A:E").format.autofitColumns();
  ws.activate();
  ws.getRange("A1").select();
  tam.delete();
  await ctx.sync();
  return { ok: true, sheet: "Bán hàng", dong: rows.length - 1 };
});
