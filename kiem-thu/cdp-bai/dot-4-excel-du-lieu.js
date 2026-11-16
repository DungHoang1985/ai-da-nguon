// Đợt 4: dựng lại dữ liệu thử trong Excel (xóa rồi tạo lại ba sheet thử; không đụng sheet khác).
//   "Tiến độ"  — bảng tiến độ 12 hạng mục (bài chuẩn hóa bảng, biểu đồ, chuẩn bị in)
//   "Mô hình"  — mô hình dòng tiền 5 năm, CÀI 3 hằng số gõ cứng (D6, E9, C12) và 1 công thức lệch hàng (E10)
//   "Hồ sơ"    — bảng theo dõi hồ sơ có 2 dòng trùng mã
// Chạy: node kiem-thu/cdp.mjs eval 9229/Excel kiem-thu/cdp-bai/dot-4-excel-du-lieu.js
return Excel.run(async (ctx) => {
  const wb = ctx.workbook;
  const names = ["Tiến độ", "Mô hình", "Hồ sơ", "Nhật ký AI"];
  wb.worksheets.load("items/name");
  wb.names.load("items/name");
  await ctx.sync();
  // Giữ ít nhất một sheet: thêm sheet tạm nếu mọi sheet đều sắp bị xóa
  const keep = wb.worksheets.items.filter((w) => !names.includes(w.name) && !/^(Pivot|Tổng hợp|AIDN_)/.test(w.name));
  if (!keep.length) wb.worksheets.add("Sheet1");
  for (const w of wb.worksheets.items) if (names.includes(w.name) || /^(Pivot|Tổng hợp)/.test(w.name)) w.delete();
  for (const n of wb.names.items) n.delete();
  await ctx.sync();

  const td = wb.worksheets.add("Tiến độ");
  const rows = [
    ["Hạng mục", "Nhóm", "Kế hoạch", "Thực tế", "Ghi chú"],
    ["San nền", "Hạ tầng", 1, 1, "Xong tháng 6"],
    ["Đường nội bộ", "Hạ tầng", 0.8, 0.62, "Chậm do mưa"],
    ["Cấp thoát nước", "Hạ tầng", 0.7, 0.55, ""],
    ["Điện chiếu sáng", "Hạ tầng", 0.5, 0.3, ""],
    ["Móng khối A", "Khối A", 0.6, 0.55, ""],
    ["Thân khối A", "Khối A", 0.35, 0.2, "Thiếu nhân công"],
    ["Hoàn thiện khối A", "Khối A", 0.1, 0, ""],
    ["Móng khối B", "Khối B", 0.4, 0.4, ""],
    ["Thân khối B", "Khối B", 0.15, 0.05, ""],
    ["Hoàn thiện khối B", "Khối B", 0, 0, "Chưa khởi công"],
    ["Cảnh quan", "Tiện ích", 0.1, 0, "Chưa khởi công"],
    ["Nhà điều hành", "Tiện ích", 0.9, 0.85, ""],
  ];
  td.getRange("A1:E13").values = rows;

  const mh = wb.worksheets.add("Mô hình");
  mh.getRange("A1:F13").formulas = [
    ["Mô hình dòng tiền dự án (tỷ đồng)", "", "", "", "", ""],
    ["Giả định", "", "", "", "", ""],
    ["Tăng giá bán mỗi năm", 0.08, "", "", "", ""],
    ["Tỷ lệ chi phí trên doanh thu", 0.62, "", "", "", ""],
    ["Chỉ tiêu", "Năm 1", "Năm 2", "Năm 3", "Năm 4", "Năm 5"],
    ["Doanh thu", 500, "=B6*(1+$B$3)", 590, "=D6*(1+$B$3)", "=E6*(1+$B$3)"],
    ["Chi phí", "=B6*$B$4", "=C6*$B$4", "=D6*$B$4", "=E6*$B$4", "=F6*$B$4"],
    ["Lợi nhuận gộp", "=B6-B7", "=C6-C7", "=D6-D7", "=E6-E7", "=F6-F7"],
    ["Chi phí quản lý", "=B6*0.05", "=C6*0.05", "=D6*0.05", 31, "=F6*0.05"],
    ["Lợi nhuận trước thuế", "=B8-B9", "=C8-C9", "=D8-D9", "=E8-E7", "=F8-F9"],
    ["Thuế", "=B10*0.2", "=C10*0.2", "=D10*0.2", "=E10*0.2", "=F10*0.2"],
    ["Lợi nhuận sau thuế", "=B10-B11", 82, "=D10-D11", "=E10-E11", "=F10-F11"],
    ["Tổng lợi nhuận sau thuế 5 năm", "=SUM(B12:F12)", "", "", "", ""],
  ];

  const hs = wb.worksheets.add("Hồ sơ");
  hs.getRange("A1:D8").values = [
    ["Mã hồ sơ", "Thủ tục", "Tình trạng", "Người phụ trách"],
    ["HS-001", "Chấp thuận chủ trương đầu tư", "Chờ", "Tổ pháp lý"],
    ["HS-002", "Quy hoạch chi tiết 1/500", "Chờ", "Tổ quy hoạch"],
    ["HS-003", "Giao đất", "Xong", "Tổ pháp lý"],
    ["HS-002", "Quy hoạch chi tiết 1/500", "Chờ", "Tổ quy hoạch"],
    ["HS-004", "Giấy phép xây dựng", "Chờ", "Tổ xây dựng"],
    ["HS-001", "Chấp thuận chủ trương đầu tư", "Chờ", "Tổ pháp lý"],
    ["HS-005", "Nghiệm thu phòng cháy", "Chưa nộp", "Tổ xây dựng"],
  ];
  td.activate();
  await ctx.sync();
  const loi = mh.getRange("B13");
  loi.load("values");
  await ctx.sync();
  return { sheets: ["Tiến độ (A1:E13)", "Mô hình (A1:F13)", "Hồ sơ (A1:D8)"], caiLoi: { hangSo: ["D6", "E9", "C12"], lechHang: ["E10"] }, tongLoiNhuan: loi.values[0][0] };
});
