## TÓM TẮT
Máy 16.0.20527 (Current Channel) đủ cho ExcelApi 1.21 (cần từ 20131) và ExcelApiDesktop 1.1 (cần từ 19231), nên hầu hết API dưới đây dùng được. Công cụ Excel hiện có (21 công cụ) còn thiếu các mảng quan trọng cho báo cáo pháp lý – bất động sản: in ấn, sửa biểu đồ và PivotTable có sẵn, tên vùng, bảo vệ sheet, làm sạch dữ liệu, nhật ký sửa đổi. Tất cả đều có API chính thức. Không có API Goal Seek, Data Table, Solver. Có thể mô phỏng bằng vòng lặp ghi, tính lại, đọc trong add-in. Excel không có Track Changes qua API, nên cần tự ghi nhật ký.

### PH1 [chac_chan] a. Biểu đồ có sẵn và getImage
KL: Sửa được biểu đồ có sẵn: chartType (1.7), title, legend, dataLabels, axes, series (add, setValues, setXAxisValues), setData(range, seriesBy) để đổi nguồn (1.1), plotBy, setPosition, đổi kích thước. Chart.getImage(width, height, fittingMode) trả PNG Base64 (1.2). Chart.getDataRange chỉ có ở preview, không dùng. Công cụ hiện có chỉ tạo mới, chưa sửa. Theo trí nhớ: đọc công thức nguồn của từng series không có.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/excel/excel.chart

### PH2 [chac_chan] b. PivotTable có sẵn và slicer
KL: Sửa được: rowHierarchies, columnHierarchies, dataHierarchies và filterHierarchies thêm/bớt (1.8); layout (1.8); refresh() (1.3); applyFilter theo nhãn, giá trị, ngày, chọn tay (1.12); refreshOnOpen (1.13); getDataSourceString và getDataSourceType (1.15). Không có hàm đổi nguồn dữ liệu, phải xóa và tạo lại. Slicer có từ 1.10, gắn với Table hoặc PivotTable. Theo trí nhớ: hàm tổng hợp đặt qua summarizeBy của dataHierarchy.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottable | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-12-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-10-requirement-set

### PH3 [chac_chan] c. In ấn (PageLayout)
KL: Worksheet.pageLayout (1.9) có đủ: setPrintArea, setPrintTitleRows và setPrintTitleColumns, orientation, paperSize (A4, A3, Legal...), zoom (scale; theo trí nhớ còn tùy chọn vừa N trang rộng/cao), setPrintMargins(đơn vị), căn giữa, in lưới và tiêu đề, thứ tự in, headersFooters (đầu/chân trang). Desktop 1.1 thêm ảnh đầu/chân trang, printQuality. Không có API mở xem trước hay gọi lệnh in.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set

### PH4 [chac_chan] d. Dấu vết soát lại
KL: Tô sáng ô bằng format.fill. Ghi chú kiểu note: notes.add (1.18). Nhận xét kiểu luồng: workbook.comments.add (1.10), kèm trả lời, resolve và nhắc tên (1.11), sự kiện thêm/sửa/xóa (1.12). Lưu ý excel_add_comment hiện dùng comments (luồng) dù mô tả ghi 'note'. Sự kiện: worksheet.onChanged (1.7), onFormulaChanged (1.13), chỉ chạy khi add-in đang mở. Không thấy API Track Changes cho Excel, nên phải tự ghi nhật ký.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-18-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-12-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-13-requirement-set

### PH5 [can_thu] e. Kiểm toán công thức
KL: Có: getDirectPrecedents (1.12) và getDirectDependents (1.13), đã dùng trong excel_trace; getSpecialCells (1.9) lọc công thức, hằng số, ô trống, lỗi; getSpillParent (1.12). Tên vùng: workbook.names (theo trí nhớ ExcelApi 1.4). Không thấy API phát hiện tham chiếu vòng. linkedWorkbooks chỉ có ở Excel web (ExcelApiOnline), không có trên desktop. Liên kết ngoài trên desktop phải quét công thức tìm '[' hoặc tên tệp .xls.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-12-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-13-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-online-requirement-set

### PH6 [chac_chan] f. Làm sạch dữ liệu và định dạng Việt Nam
KL: Có (1.9): removeDuplicates(columns, includesHeader), autoFill, flashFill, replaceAll(text, replacement, criteria); sort (1.2); autoFilter (1.9). numberFormatLocal (1.7) và formulasLocal (1.1) theo ngôn ngữ người dùng. Application.cultureInfo, decimalSeparator, thousandsSeparator (1.11) và datetimeFormat (1.12) cho biết dấu phân cách và mẫu ngày của máy. numberFormat 'dd/mm/yyyy' dùng được vì là mã không phụ thuộc locale.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-9-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-11-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-12-requirement-set

### PH7 [can_thu] g1. Ảnh vùng, bảo vệ, nhóm, chép sheet, PDF, hàm tùy biến
KL: Range.getImage() trả PNG Base64 (1.7). Bảo vệ sheet: worksheet.protection và range.format.protection.locked (theo trí nhớ 1.2); AllowEditRange chỉ có ở Excel web. Nhóm hàng/cột: Range.group (1.10). Chép sheet trong tệp: Worksheet.copy (1.7). Chèn sheet từ tệp khác: insertWorksheetsFromBase64 (1.13). PDF: Document.getFileAsync(FileType.Pdf) có trong enum chung, chưa thấy tài liệu xác nhận riêng cho Excel. Hàm tùy biến cần functions.json, thẻ Runtimes (lifetime long) và SharedRuntime 1.1; chạy được trên desktop, nhưng công sức lớn.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-7-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-13-requirement-set | https://learn.microsoft.com/en-us/javascript/api/office/office.filetype | https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime

### PH8 [khong_lam_duoc] g2. Goal Seek, Data Table, Solver
KL: Không có API. Các bảng API của ExcelApi 1.18 đến 1.21, ExcelApiDesktop 1.1, ExcelApiOnline và danh sách preview đều không liệt kê. Đã đọc nửa đầu trang Excel.Range (khoảng 100.000/203.000 ký tự), cũng không thấy. Cách thay thế: add-in tự chạy vòng ghi giá trị, tính lại (Application.calculate), đọc kết quả, rồi khôi phục giá trị gốc. Dùng chia đôi cho Goal Seek, và lưới kịch bản cho bảng nhạy cảm.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-preview-apis | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set

## ĐỀ XUẤT TÍNH NĂNG
1. excel_page_setup — khả thi=chac_chan giá trị=cao công=nho hiện có=chua
   Cài đặt in cho báo cáo lãnh đạo: vùng in, tiêu đề lặp, A4 ngang, co vừa trang, lề, đầu/chân trang (tên dự án, số trang).
   API: Worksheet.pageLayout (ExcelApi 1.9) (ExcelApi 1.9)
2. excel_edit_chart — khả thi=chac_chan giá trị=cao công=vua hiện có=mot_phan
   Sửa biểu đồ có sẵn: đổi kiểu, tiêu đề, chuỗi, trục, nhãn, nguồn dữ liệu, vị trí. Liệt kê biểu đồ trong sheet.
   API: Chart, ChartSeries, ChartAxes, Chart.setData (ExcelApi 1.1 đến 1.8) (ExcelApi 1.8)
3. excel_named_ranges — khả thi=chac_chan giá trị=cao công=nho hiện có=chua
   Liệt kê, tạo, sửa tên vùng cho giả định mô hình tài chính (lãi suất, đơn giá, tỷ lệ), giúp công thức dễ đọc.
   API: Workbook.names (theo trí nhớ: ExcelApi 1.4) (ExcelApi 1.4 (theo trí nhớ))
4. excel_clean_data — khả thi=chac_chan giá trị=cao công=nho hiện có=chua
   Làm sạch bảng hồ sơ, hợp đồng: xóa trùng, tìm và thay hàng loạt, tự điền chuỗi ngày hạn, Flash Fill. Hỏi trước khi xóa.
   API: Range.removeDuplicates, replaceAll, autoFill, flashFill (ExcelApi 1.9) (ExcelApi 1.9)
5. excel_audit_model — khả thi=can_thu giá trị=cao công=vua hiện có=mot_phan
   Kiểm toán mô hình: hằng số cứng lẫn trong vùng công thức, công thức lệch hàng, liên kết ngoài, tham chiếu vòng (thử), tên lỗi.
   API: Range.getSpecialCells (1.9), formulas, getDirectPrecedents (1.12) (ExcelApi 1.12)
6. excel_change_log — khả thi=chac_chan giá trị=cao công=vua hiện có=chua
   Nhật ký thay đổi do AI: sheet 'Nhật ký AI' ghi ô, giá trị cũ, mới, lý do; tô sáng ô đã sửa. Thay cho Track Changes Excel không có.
   API: Range.values, format.fill, Worksheet.onChanged (1.7) (ExcelApi 1.7)
7. excel_goal_seek_sim — khả thi=can_thu giá trị=cao công=vua hiện có=chua
   Giải ngược (Goal Seek giả lập): tìm đầu vào để đạt mục tiêu (IRR, lợi nhuận). Chia đôi bằng vòng ghi, tính lại, đọc; khôi phục nếu hủy.
   API: Range.values, Application.calculate (không có API Goal Seek) (ExcelApi 1.1)
8. excel_sensitivity_table — khả thi=can_thu giá trị=cao công=vua hiện có=chua
   Bảng nhạy cảm 1-2 biến cho mô hình dự án: thử lưới giá trị, ghi kết quả ra bảng tĩnh, khôi phục giá trị gốc. Thay Data Table.
   API: Range.values, Application.calculate (không có API Data Table) (ExcelApi 1.1)
9. excel_protect_sheet — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=chua
   Khóa ô công thức, chừa ô nhập; bảo vệ sheet với tùy chọn cho phép lọc, sắp xếp. Mật khẩu do người dùng nhập, không lưu.
   API: Worksheet.protection, Range.format.protection (theo trí nhớ: 1.2) (ExcelApi 1.2 (theo trí nhớ))
10. excel_pivot_edit — khả thi=chac_chan giá trị=trung_binh công=vua hiện có=mot_phan
   Sửa PivotTable có sẵn: thêm/bớt trường, đổi hàm tổng hợp, lọc, làm mới; thêm slicer cho bảng theo dõi.
   API: PivotTable hierarchies (1.8), PivotField.applyFilter (1.12), Slicer (1.10) (ExcelApi 1.12)
11. excel_group_outline — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=chua
   Nhóm hàng/cột thành dàn ý thu gọn để báo cáo lãnh đạo xem tóm tắt, mở chi tiết khi cần.
   API: Range.group, ungroup, Worksheet.showOutlineLevels (ExcelApi 1.10) (ExcelApi 1.10)
12. excel_get_overview_plus — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=mot_phan
   Mở rộng tổng quan: dấu phân cách và mẫu ngày của máy, danh sách tên vùng, vùng in, trạng thái bảo vệ, biểu đồ, pivot, liên kết ngoài.
   API: Application.cultureInfo (1.11, 1.12), names, pageLayout, protection (ExcelApi 1.12)
13. excel_comment_note_modes — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=mot_phan
   Tách rõ nhận xét luồng và ghi chú (note); thêm trả lời, đánh dấu đã giải quyết. Sửa mô tả excel_add_comment cho đúng.
   API: workbook.comments (1.10, 1.11), workbook.notes (1.18) (ExcelApi 1.18)
14. excel_view_image — khả thi=can_thu giá trị=trung_binh công=vua hiện có=chua
   Chụp ảnh vùng ô hoặc biểu đồ cho model có hỗ trợ ảnh nhìn lại bố cục, kiểm tra định dạng. Cần sửa llm.ts để gửi ảnh trong kết quả công cụ.
   API: Range.getImage (1.7), Chart.getImage (1.2) (ExcelApi 1.7, model hỗ trợ ảnh)
15. excel_insert_sheets_from_file — khả thi=chac_chan giá trị=trung_binh công=vua hiện có=chua
   Chèn sheet từ tệp .xlsx người dùng đính kèm (biểu mẫu tiến độ, mô hình mẫu) vào tệp đang mở. Khuyến nghị chọn sheet cần chèn.
   API: Workbook.insertWorksheetsFromBase64 (ExcelApi 1.13) (ExcelApi 1.13)
16. excel_export_pdf — khả thi=can_thu giá trị=thap công=vua hiện có=chua
   Xuất tệp thành PDF để gửi lãnh đạo. Chưa xác nhận Excel hỗ trợ, và WebView2 phải có cách lưu tệp.
   API: Office.context.document.getFileAsync(Office.FileType.Pdf) (Cần thử trên Excel 16.0.20527)

## KHÔNG XÁC MINH ĐƯỢC
- Excel có hỗ trợ getFileAsync với FileType.Pdf hay không (tài liệu chỉ có ví dụ Word, chưa thấy ghi cho Excel)
- Có API phát hiện tham chiếu vòng hay không (không thấy trong các bảng đã đọc; chưa kiểm tra từng lớp)
- Tên thuộc tính co vừa trang của PageLayoutZoomOptions, cấu trúc HeaderFooterGroup, thuộc tính summarizeBy của dataHierarchy, tập API của names và protection: theo trí nhớ, chưa mở trang riêng
- Trang Excel.Range chỉ đọc nửa đầu (khoảng 100.000/203.000 ký tự); các bảng ExcelApi 1.14 đến 1.17 chưa đọc
- Chart.add có tạo được các kiểu mới (Waterfall, Funnel, Treemap) hay không, dù enum ChartType có liệt kê
- Cấu hình hàm tùy biến với manifest.xml dạng add-in only: tài liệu đã đọc chủ yếu nói về unified manifest và shared runtime