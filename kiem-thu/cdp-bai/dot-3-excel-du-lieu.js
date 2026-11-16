// Đợt 3 (bài Chạy thật B-1, B-2): dựng bảng tiến độ thử trong Excel (sheet "Tiến độ").
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-bai/dot-3-excel-du-lieu.js
return Excel.run(async (ctx) => {
  let sh = ctx.workbook.worksheets.getItemOrNullObject("Tiến độ");
  await ctx.sync();
  if (sh.isNullObject) sh = ctx.workbook.worksheets.add("Tiến độ");
  sh.activate();
  const rows = [
    ["Hạng mục", "Kế hoạch (%)", "Thực tế (%)", "Ghi chú"],
    ["San nền", 100, 100, "Xong tháng 6"],
    ["Hạ tầng kỹ thuật", 80, 62, "Chậm do mưa"],
    ["Móng khối A", 60, 55, ""],
    ["Thân khối A", 35, 20, "Thiếu nhân công"],
    ["Cảnh quan", 10, 0, "Chưa khởi công"],
  ];
  const r = sh.getRange("A1:D6");
  r.values = rows;
  sh.getRange("A1:D1").format.font.bold = true;
  r.format.autofitColumns();
  await ctx.sync();
  return { sheet: "Tiến độ", hang: rows.length, tongKeHoach: 285, tongThucTe: 237 };
});
