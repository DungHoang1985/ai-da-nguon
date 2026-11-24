## TÓM TẮT
Office.js PowerPoint bản chính thức mới nhất là PowerPointApi 1.10 (cần Windows từ build 19610; máy 16.0.20527 đủ). Làm được chắc chắn: chụp ảnh slide cho model nhìn, đọc layout/placeholder/bảng màu, định dạng chữ từng đoạn, bảng đầy đủ (gộp ô), nhóm/lớp, đọc và đặt lựa chọn, alt text, siêu liên kết. Không có API cho biểu đồ, SmartArt, ghi chú diễn giả, ẩn slide, chuyển cảnh và phông chủ đề. Riêng chèn ảnh (addPicture) mới chỉ ở bản preview. Các phần thiếu chỉ làm được bằng cách xuất slide thành .pptx, sửa XML rồi chèn lại, tức là phải thử bằng Tự kiểm tra. Ưu tiên làm: ảnh slide cho model, kiểm bài theo mẫu, đọc chủ đề/layout, định dạng đoạn chữ, đọc lựa chọn hiện tại.

### PH1 [chac_chan] a. Master, layout, placeholder, chủ đề
KL: Đọc được: slideMasters và layouts (1.3), shapes của master/layout (1.3), SlideLayout.type và placeholderFormat.type/containedType (1.8). Bảng 12 màu chủ đề đọc và ghi được qua ThemeColorScheme.getThemeColor/setThemeColor trên slide, layout, master (1.10). Nền slide/layout/master có trong 1.10. KHÔNG có API đọc phông chủ đề: chỉ đọc được bằng cách giải nén theme1.xml từ slide.exportAsBase64 (repo đã dùng JSZip đọc sldSz). Áp bố cục: Slide.applyLayout (1.8), thêm slide theo layoutId/slideMasterId (1.3). Kích thước slide có PageSetup.slideWidth/slideHeight (1.10), thay được cách đọc XML hiện nay.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-1-10-requirement-set | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-1-8-requirement-set | https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slidemaster?view=powerpoint-js-preview

### PH2 [chac_chan] b. Ảnh chụp slide cho model
KL: Slide.getImageAsBase64(options) có từ PowerPointApi 1.8, trả PNG base64. Tùy chọn width/height (pixel): bỏ cả hai thì dùng kích thước thật; chỉ cho một chiều thì giữ tỷ lệ. Shape.getImageAsBase64 (1.10) có thêm format, width, height để chụp từng hình. Bản 1.8 cần Windows từ build 18730; máy 16.0.20527 vượt xa. Chưa kiểm được lớp gọi model của repo có nhận ảnh trong kết quả công cụ hay không.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slide?view=powerpoint-js-preview | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-1-8-requirement-set

### PH3 [chac_chan] c. Hình, chữ, bảng, ảnh, liên kết, alt text
KL: Có: addTextBox, addGeometricShape (hàng trăm kiểu), addLine (1.4); addTable, addGroup, setZOrder, zOrderPosition (1.8); rotation, visible, altTextTitle/Description, isDecorative (1.10). TextFrame: autoSizeSetting, wordWrap, lề, canh dọc (1.4). TextRange: font, getSubstring, bullet visible và canh lề (1.4); bullet style/type, indentLevel (1.10). Bảng: sửa ô, gộp ô mergeCells, tô màu, viền, thêm/xóa hàng cột (1.8–1.9). Siêu liên kết: đọc từ 1.5/1.6, thêm bằng setHyperlink (1.10), chỉ URL/tệp/mailto. Chèn ảnh addPicture CHỈ có ở preview, không dùng cho bản chạy; đường vòng ShapeFill.setImage (1.8).
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection?view=powerpoint-js-preview | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-1-9-requirement-set | https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textframe?view=powerpoint-js-preview | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-preview-apis

### PH4 [can_thu] d. Biểu đồ gốc và SmartArt
KL: KHÔNG có API tạo hay sửa biểu đồ. ShapeType chỉ có Chart, SmartArt, Diagram để nhận diện; lớp Graphic trong preview chỉ có convertToShape. Vì vậy cách repo đang làm (pptxgenjs rồi insertSlidesFromBase64) là đúng hướng. Cập nhật số liệu biểu đồ đã có chỉ làm được bằng cách xuất slide .pptx, sửa chartN.xml và sổ tính nhúng rồi chèn lại và xóa slide cũ (theo suy luận từ định dạng OOXML, cần thử). SmartArt: không tạo, không sửa được; thay bằng ppt_add_process_flow.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapetype?view=powerpoint-js-preview | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-preview-apis

### PH5 [chac_chan] e. Ghi chú diễn giả, ẩn slide, chuyển cảnh, chèn slide, tags, PDF
KL: Lớp Slide và Presentation (kể cả preview, cập nhật 23/10/2026) KHÔNG có ghi chú diễn giả, ẩn slide, chuyển cảnh hay hiệu ứng; chỉ làm được bằng sửa XML trong tệp xuất rồi chèn lại (cần thử). Có: insertSlidesFromBase64 (1.2, formatting UseDestinationTheme/KeepSourceFormatting, targetSlideId); exportAsBase64 (1.8); tags trên presentation/slide/shape (1.3). Xuất PDF: Office.context.document.getFileAsync(Office.FileType.Pdf) chạy trên Windows, Mac, iPad, không chạy trên web; mỗi lát tối đa 4 MB.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slide?view=powerpoint-js-preview | https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation?view=powerpoint-js-preview | https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=powerpoint-js-preview

### PH6 [chac_chan] f. Đọc và đặt lựa chọn hiện tại
KL: Có từ PowerPointApi 1.5: getSelectedSlides, getSelectedShapes, getSelectedTextRange (và bản OrNullObject), setSelectedSlides, Slide.setSelectedShapes, TextRange.setSelected. Slide đang hiển thị: getActiveSlideOrNullObject và sự kiện onSlideSelectionChanged mới ở preview. Repo đã dùng supports("1.5") để đọc slide đang chọn trong ppt_get_overview, nhưng chưa có công cụ đọc hình hay đoạn chữ đang bôi đen.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation?view=powerpoint-js-preview

### PH7 [chac_chan] g. Bộ yêu cầu và máy hiện tại
KL: Bản chính thức mới nhất trên trang yêu cầu là PowerPointApi 1.10 (Windows tối thiểu 2601, build 19610.20002). Chưa thấy 1.11: trang yêu cầu lưu ở bản chụp 12/2025, danh sách preview cập nhật 23/10/2026 không có số 1.11. Máy 16.0.20527 vượt ngưỡng nên 1.10 chạy được, nhưng Office bản bán buôn/2021 không có 1.8 trở lên. Luôn kiểm tra isSetSupported('PowerPointApi','1.10') bằng chuỗi, không dùng số.
NGUỒN: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets | https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-preview-apis

### PH8 [chac_chan] h. Khoảng trống của 14 công cụ hiện có
KL: Đọc pptTools.ts và pptAdvanced.ts: 14 công cụ (overview, add_slide, set_shape_text, add_textbox, add_shape, add_table, go_to_slide, add_chart, add_process_flow, move/duplicate/delete_slide, update_shape, delete_shape). Thiếu: chụp ảnh slide; đọc chủ đề/màu/phông/placeholder; đọc font-cỡ-màu theo đoạn; sửa bảng đã có; ghi chú diễn giả; sửa số liệu biểu đồ có sẵn (add_chart chỉ tạo slide mới); chèn ảnh; kiểm bài theo mẫu; áp lại bố cục.
NGUỒN: D:/ADD IN (repo-main/src/office/pptTools.ts, pptAdvanced.ts)

## ĐỀ XUẤT TÍNH NĂNG
1. ppt_get_slide_image — khả thi=chac_chan giá trị=cao công=vua hiện có=chua
   Chụp slide thành PNG (rộng 960 px) đưa cho model nhìn, để tự kiểm bố cục, chữ tràn, hình chồng nhau sau mỗi lần dựng slide hoặc sơ đồ.
   API: Slide.getImageAsBase64({width}) (PowerPointApi 1.8); Shape.getImageAsBase64 (1.10) (PowerPointApi 1.8; model có khả năng nhìn ảnh; llm.ts phải nhận ảnh trong kết quả công cụ (chưa kiểm))
2. ppt_audit_deck — khả thi=chac_chan giá trị=cao công=vua hiện có=chua
   Kiểm bài theo mẫu: phông/cỡ/màu lệch chuẩn chủ đề, slide thiếu tiêu đề, chữ dưới 12 pt, hình tràn khỏi slide, chữ tràn khung, hình chồng nhau. Có chế độ tự sửa.
   API: Shape.textFrame.textRange.font, placeholderFormat.type, TextFrame.autoSizeSetting, PageSetup.slideWidth/Height, ThemeColorScheme.getThemeColor (PowerPointApi 1.10 (tối thiểu 1.8); phông chuẩn lấy từ ppt_get_theme)
3. ppt_get_theme — khả thi=chac_chan giá trị=cao công=vua hiện có=mot_phan
   Đọc chủ đề: 12 màu, phông tiêu đề/nội dung, tên master và layout, placeholder từng layout (loại, vị trí). Để model dùng đúng màu, phông, bố cục của mẫu công ty.
   API: slideMasters/layouts/shapes, placeholderFormat.type, themeColorScheme.getThemeColor (1.10); phông đọc từ theme1.xml trong exportAsBase64 qua JSZip (PowerPointApi 1.10 cho màu; 1.8 cho xuất slide; JSZip đã có trong repo)
4. ppt_format_text — khả thi=chac_chan giá trị=cao công=vua hiện có=mot_phan
   Định dạng chữ theo đoạn hoặc đoạn con: bullet bật/tắt, kiểu, cấp thụt lề, cỡ, đậm, màu, canh lề; đặt autofit, xuống dòng, lề khung chữ. Có đọc lại định dạng hiện tại.
   API: TextRange.getSubstring, font, paragraphFormat.bulletFormat/horizontalAlignment/indentLevel, TextFrame.autoSizeSetting/wordWrap/margins (PowerPointApi 1.4; bullet style/type và indentLevel cần 1.10)
5. ppt_get_selection / ppt_select — khả thi=chac_chan giá trị=cao công=nho hiện có=mot_phan
   Đọc slide, hình, đoạn chữ người dùng đang chọn để sửa đúng chỗ họ trỏ; đặt lựa chọn sau khi sửa để người dùng thấy ngay kết quả.
   API: getSelectedSlides, getSelectedShapes, getSelectedTextRangeOrNullObject, setSelectedSlides, Slide.setSelectedShapes, TextRange.setSelected (PowerPointApi 1.5)
6. ppt_change_layout — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=chua
   Đổi bố cục của slide về layout đúng mẫu công ty (Title Only, Two Content...) để sửa slide lệch mẫu; liệt kê layout và placeholder trước khi áp.
   API: Slide.applyLayout(layout), Slide.layout, slideMasters.layouts (PowerPointApi 1.8; nên hoàn tác bằng exportAsBase64 như repo đang làm)
7. ppt_edit_table — khả thi=chac_chan giá trị=cao công=vua hiện có=mot_phan
   Sửa bảng có sẵn: ghi ô, tô màu, phông, viền, gộp/tách ô, thêm/xóa hàng cột, độ rộng cột, kiểu bảng. Dùng cho bảng tiến độ, so sánh phương án, danh mục thủ tục.
   API: Shape.getTable, Table.values/mergeCells/rows/columns, TableCell.fill/font/borders/split/resize, TableAddOptions.style (PowerPointApi 1.8 (đọc, thêm) và 1.9 (định dạng, gộp ô, hàng cột))
8. ppt_insert_from_template — khả thi=chac_chan giá trị=cao công=vua hiện có=mot_phan
   Chèn slide chọn lọc từ tệp .pptx mẫu đính kèm (bìa, mục lục, slide kết) dùng theme của bài hiện tại hoặc giữ nguyên theme nguồn.
   API: Presentation.insertSlidesFromBase64(base64, {formatting, targetSlideId, sourceSlideIds}) (PowerPointApi 1.2; add-in nhận tệp đính kèm (attachments.ts); sourceSlideIds theo trí nhớ, cần thử)
9. ppt_arrange_shapes — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=mot_phan
   Nhóm/bỏ nhóm, đổi thứ tự lớp, canh đều và dàn đều khoảng cách (tính từ left/top/width), xoay, ẩn/hiện hình. Dọn bố cục khi audit báo lệch.
   API: ShapeCollection.addGroup, ShapeGroup.ungroup, Shape.setZOrder/rotation/visible/left/top (PowerPointApi 1.8; rotation và visible cần 1.10)
10. ppt_update_chart_data — khả thi=can_thu giá trị=cao công=lon hiện có=mot_phan
   Cập nhật số liệu biểu đồ gốc đã có (báo cáo tháng sau): xuất slide, sửa chart XML và sổ tính nhúng, chèn lại đúng vị trí, xóa slide cũ. Dữ liệu lấy từ Excel đính kèm.
   API: Không có API biểu đồ; Slide.exportAsBase64 + JSZip sửa ppt/charts/chartN.xml và embeddings + insertSlidesFromBase64 (PowerPointApi 1.8; phải có bước Tự kiểm tra: chèn lại còn Edit Data được; theo suy luận từ OOXML)
11. ppt_add_image — khả thi=can_thu giá trị=cao công=vua hiện có=chua
   Chèn ảnh base64 (bản đồ quy hoạch, ảnh hiện trạng, logo) vào slide đúng vị trí, kích thước, giữ tỷ lệ, kèm alt text. Dùng đường vòng vì addPicture chưa phát hành.
   API: addGeometricShape + ShapeFill.setImage (1.8); hoặc dựng slide bằng pptxgenjs rồi insertSlidesFromBase64; addPicture chỉ ở preview (PowerPointApi 1.8; cần đọc kích thước ảnh để giữ tỷ lệ; đổi sang addPicture khi thành bản chính thức)
12. ppt_set_notes — khả thi=can_thu giá trị=trung_binh công=lon hiện có=chua
   Đọc và ghi ghi chú diễn giả cho từng slide (kịch bản trình bày với lãnh đạo, căn cứ pháp lý đi kèm) bằng sửa notesSlide trong tệp xuất rồi chèn lại slide.
   API: Không có API ghi chú; Slide.exportAsBase64 + sửa ppt/notesSlides/*.xml + insertSlidesFromBase64 (PowerPointApi 1.8; chưa rõ chèn lại có giữ ghi chú không (cần Tự kiểm tra); đọc ghi chú dễ hơn ghi)
13. ppt_export_pdf — khả thi=can_thu giá trị=trung_binh công=vua hiện có=chua
   Xuất toàn bộ bài thành PDF để gửi lãnh đạo hoặc kèm tờ trình, đọc theo từng lát 4 MB rồi ghép và lưu qua hộp thoại tải xuống.
   API: Office.context.document.getFileAsync(Office.FileType.Pdf) + getSliceAsync (Windows, Mac, iPad (không chạy trên web); chưa kiểm được việc tải tệp từ khung WebView2 của add-in)
14. ppt_set_theme_style — khả thi=chac_chan giá trị=trung_binh công=nho hiện có=chua
   Đặt màu chủ đề theo màu công ty và nền slide (đơn sắc, chuyển màu, ảnh). Thao tác đổi toàn bài nên luôn hỏi xác nhận trước khi ghi.
   API: ThemeColorScheme.setThemeColor; Slide.background.fill.setSolidFill/setGradientFill/setPictureOrTextureFill (PowerPointApi 1.10; không có API đổi phông chủ đề)
15. ppt_set_alt_and_link — khả thi=chac_chan giá trị=thap công=nho hiện có=chua
   Đặt alt text, đánh dấu hình trang trí, gắn siêu liên kết (URL, tệp, email) cho hình hoặc đoạn chữ, ví dụ link văn bản pháp luật trong phụ lục.
   API: Shape.altTextTitle/altTextDescription/isDecorative; Shape.setHyperlink, TextRange.setHyperlink, HyperlinkAddOptions (PowerPointApi 1.10; không liên kết nhảy tới slide khác theo danh sách API đã đọc)
16. ppt_set_slide_flags — khả thi=can_thu giá trị=thap công=lon hiện có=chua
   Ẩn/hiện slide và đặt chuyển cảnh bằng sửa XML (show=0, p:transition) trong tệp xuất rồi chèn lại. Không có API trực tiếp; chỉ làm khi người dùng thật sự cần.
   API: Không có API; Slide.exportAsBase64 + sửa ppt/slides/slideN.xml + insertSlidesFromBase64 (PowerPointApi 1.8; theo suy luận OOXML, cần Tự kiểm tra; giá trị thấp cho việc pháp lý)

## KHÔNG XÁC MINH ĐƯỢC
- Việc insertSlidesFromBase64 giữ lại ghi chú diễn giả, biểu đồ và sổ tính nhúng khi chèn lại slide đã sửa XML (chưa thử; cần Tự kiểm tra trong Office).
- InsertSlideOptions.sourceSlideIds: chỉ biết theo trí nhớ, chưa đọc trang lớp này.
- Có PowerPointApi 1.11 hay không: trang yêu cầu cache ở bản 12/2025, danh sách preview cập nhật 23/10/2026 không thấy 1.11.
- Lớp TextRange đầy đủ (cách lấy từng đoạn) chưa đọc trang riêng; chỉ thấy getSubstring, font, paragraphFormat qua ví dụ.
- llm.ts của repo có nhận ảnh trong kết quả công cụ cho cả 3 kiểu API (OpenAI, Anthropic, Gemini) hay không: chưa đọc.
- Tải tệp PDF từ khung WebView2 của add-in trên máy này: chưa kiểm.