# AI Đa Nguồn — add-in AI cho Word, Excel, PowerPoint

Add-in Office dùng API key của nhiều nhà cung cấp; AI tự thao tác trên tài liệu như add-in Claude chính thức. Dùng cá nhân (anh Dũng, pháp lý – bất động sản). Phiên bản đang chạy và lịch sử phiên bản: đầu `tasks/HANDOFF.md` (số gốc: `version` trong `package.json`).

## Đọc khi mở phiên

1. `tasks/HANDOFF.md` — trạng thái, quyết định, vấn đề mở, bước tiếp theo, bài học. **Chỉ bắt buộc tệp này.**
2. Trước khi sửa mã của phần nào, đọc quy tắc của phần đó ("không làm trái", đã thăm dò trong Office):

| Sửa gì | Đọc |
|---|---|
| Công cụ Word, `outline.ts`, `worddiff.ts`, `toolgroups.ts`, `claims.ts` | `docs/quy-tac/word.md` |
| Công cụ Excel, `auditlog.ts`, `modelaudit.ts`, `formularisk.ts`, `datetext.ts` | `docs/quy-tac/excel.md` |
| Công cụ PowerPoint, `tablefit.ts`, `deckcheck.ts` | `docs/quy-tac/powerpoint.md` |
| Nguồn AI, chế độ Miễn phí / Trả phí, đăng nhập, đo token, model mặc định | `docs/quy-tac/nguon-ai.md` |
| Vòng tác tử, tệp đính kèm, bộ nhớ dài hạn, ghi chéo ứng dụng, kỹ năng | `docs/quy-tac/tac-tu.md` |
| Kiểm trong Office | `docs/kiem-thu-office.md` |
| Đụng tới Office.js, WebView2, bài thử, phát hành | `docs/bai-hoc.md` (tìm theo từ khóa, không cần đọc hết) |

3. Chỉ đọc khi cần: kế hoạch gốc `tasks/ke-hoach-v4.md`; chi tiết từng đợt `tasks/dot/dot-<N>/dot-<N>.md` (bằng chứng ở `bang-chung/` cùng thư mục); bản vá và buổi thử lẻ `tasks/viec-le/`; bảng so sánh với bản chính thức `docs/so-sanh-claude-chinh-thuc.md`; tài liệu cũ `tasks/luu/`.

## Nguồn mã

- `D:\ADD IN` là bản làm việc của nhánh `main`. Kho công khai: https://github.com/DungHoang1985/ai-da-nguon — `main` (mã), `gh-pages` (bản đang chạy: https://dunghoang1985.github.io/ai-da-nguon/).
- Trước khi sửa: `git status` sạch và `git pull`.
- Đẩy lên `main` là GitHub Actions (`.github/workflows/phat-hanh.yml`) kiểm thử, kiểm kiểu, đóng gói rồi phát hành lên `gh-pages`; trượt thì không phát hành. Nhánh `thu/**` chạy cùng quy trình nhưng không đẩy.

## Lệnh

```
npm ci                      # .npmrc phải trỏ registry.npmjs.org
npm test                    # mocha, mọi kiểm thử phải đạt (1.044 ngày 26/11/2026)
npx tsc --noEmit -p .       # không lỗi kiểu
npm run lint                # 0 lỗi (cảnh báo Office.js được phép); CI chặn khi có lỗi
npm run build               # ra dist/, manifest tự đổi sang địa chỉ GitHub Pages
```

Thử bản chưa phát hành trong Office: `npm run build` rồi `node kiem-thu/cdp.mjs serve 9229 dist` (chạy nền). Chi tiết: `docs/kiem-thu-office.md`.

## Cấu trúc

```
src/core/                lõi không phụ thuộc Office, kiểm thử bằng npm test
  agent/                 turn.ts (điều phối một lượt; giao diện chỉ gọi runTurn) · loop.ts (vòng tác tử, lịch sử chỉ thêm)
                         llm.ts (gom 3 kiểu API OpenAI / Anthropic / Gemini, thử lại, bộ nhớ đệm, khối suy nghĩ)
                         prompt.ts · compaction.ts (tự tóm tắt) · history.ts (chuẩn hóa lịch sử khi đổi nguồn)
                         rotation.ts (xoay vòng nguồn miễn phí) · thinking.ts (Suy nghĩ kỹ) · types.ts
  nguồn, model, chế độ   providers.ts (dialect, keyFrom, nguồn tự khai) · models.ts · catalog.ts · modes.ts
                         toolgroups.ts (nạp công cụ theo việc) · inputsize.ts (đo token) · providerCheck.ts
                         oauth.ts · credits.ts · errors.ts (maskSecrets) · sse.ts
  Word                   outline.ts (mục lục điều khoản) · worddiff.ts (sửa đoạn, Hoàn tác) · consistency.ts
                         compare.ts · pageformat.ts · formfill.ts · adminDoc.ts (văn bản hành chính, hợp đồng)
                         docxparas.ts (tệp .docx đính kèm) · claims.ts · textcheck.ts · legalcheck.ts
  Excel                  auditlog.ts · formularisk.ts · modelaudit.ts · goalseek.ts · datetext.ts · table.ts
  PowerPoint             deckcheck.ts · tablefit.ts · imagefit.ts
  chung                  memory.ts · skills.ts · skillpack.ts (kỹ năng dựng sẵn) · crossapp.ts (ghi chéo)
                         attachedFiles.ts · attachments.ts · pdfPages.ts · pdfRender.ts
                         conversations.ts · db.ts · storage.ts · backup.ts · share.ts · refs.ts · keys.ts · version.ts
src/office/              công cụ Office.js
  Word                   wordTools.ts · wordReview.ts · wordAdmin.ts · wordForm.ts · selftestWord.ts
  Excel                  excelTools.ts · excelCore.ts · excelAdvanced.ts · excelModel.ts · excelTrail.ts (bọc nhật ký) · selftestExcel.ts
  PowerPoint             pptTools.ts · pptCore.ts · pptAdvanced.ts · pptDesign.ts
  chung                  tools.ts (chọn bộ công cụ theo ứng dụng) · common.ts · crossExec.ts · host.ts · navigate.ts
                         dialog.ts · util.ts · selftest.ts · selftestPdf.ts
src/taskpane/components/ App.tsx · Chat.tsx · Settings.tsx · ModelPicker.tsx · SignInCard.tsx · ProviderCheck.tsx · BackupCard.tsx
src/dialog/              dang-nhap.html/.ts — trang hộp thoại đăng nhập
test/unit/               mocha (dữ liệu mẫu ở test/fixtures/)
kiem-thu/                kiểm trong Office bằng mã (cdp.mjs, cdp-bai/, tệp thử, desktop ẩn)
cai-dat/                 tệp .cmd cài và gỡ add-in (registry WEF\Developer)
docs/                    quy-tac/ (quy tắc theo phần), bai-hoc.md, kiem-thu-office.md, so-sanh-claude-chinh-thuc.md
tasks/                   HANDOFF.md · ke-hoach-v4.md · dot/ · viec-le/ · luu/
```

## Quy tắc làm việc với anh Dũng

- **Không chiếm màn hình của anh**: kiểm thử hay thao tác phải bật Office, cửa sổ thì chạy không cửa sổ hoặc trên desktop ẩn (`kiem-thu/mo-office.ps1`); không dùng công cụ điều khiển màn hình khi chưa hỏi; không có cách tách thì dừng và hỏi.
- Trả lời tiếng Việt, kết luận trước, giải thích ngắn. Không kể lại các bước đã làm.
- **Anh không gõ lệnh, không cài tay.** Việc buộc phải có anh (đăng nhập, bấm xác nhận) thì gom lại, nói rõ từng cú bấm.
- Câu hỏi cần anh duyệt: ô tích chọn, 2–4 phương án, khuyến nghị đặt đầu kèm "(Khuyến nghị)", tối đa 4 câu. Điều tự kiểm được thì không hỏi.
- Trước khi làm: nêu tiêu chí nghiệm thu. Khi giao: bảng Tiêu chí · Kết quả (Đạt / Không đạt / Chưa kiểm được) · Bằng chứng. Không có bằng chứng thì không ghi Đạt.
- Bước không đạt: làm lại tối đa 3 lần, rồi dừng và báo.
- Cuối phiên: cập nhật `tasks/HANDOFF.md` (đã xong · đang làm · quyết định · vấn đề mở · bước tiếp theo · bài học), giữ dưới 15 KB — chi tiết cũ chuyển vào `tasks/luu/`, chi tiết đợt vào tệp đợt. Anh sửa kết quả hoặc lỗi lặp lại: ghi vào đầu `docs/bai-hoc.md`.

## Quy tắc kỹ thuật chung

- "Đạt" nghĩa là AI **thao tác thật** trên tài liệu trong Office. Ba mức bằng chứng, ghi đúng tên: **Tự kiểm tra** (nút trong add-in) · **Chạy thật** (giao việc cho model thật) · **Kiểm thử máy** (`npm test`).
- Thêm hoặc sửa công cụ Office: thêm bước vào `src/office/selftest*.ts`.
- Thao tác rủi ro (ghi đè ô có dữ liệu, xóa ô/đoạn/slide/sheet, chấp nhận hoặc từ chối tất cả sửa đổi) luôn hỏi người dùng kèm lý do. Word luôn ghi Track Changes.
- Lịch sử hội thoại gửi model là **chỉ thêm**: không sửa, không cắt tin nhắn cũ giữa hai lần tóm tắt; lời nhắc hệ thống không chứa thứ hay đổi (ngày giờ, trạng thái tài liệu). Sửa khác đi là mất bộ nhớ đệm (tốn tiền) — kiểm thử chặn ở `test/unit/dot0b.test.ts`.
- Tài liệu nhà cung cấp và Office.js đổi nhanh: đối chiếu tài liệu chính thức hiện hành trước khi viết mã gọi API.
- Không thêm máy chủ riêng, không đổi khung giao diện, không viết lại từ đầu.
- Kho GitHub **công khai**: không đưa key, token, thông tin riêng tư vào bất kỳ tệp nào; câu lỗi và bằng chứng qua `maskSecrets`.
- Tài khoản của anh: không đăng nhập thay, không tạo hay thu hồi key, không nhận tín dụng, không bấm đồng ý điều khoản, không xóa key thật đang lưu khi chưa hỏi. Gom vào mục "Việc buộc phải có anh ở máy" của tệp đợt.
- Dữ liệu add-in (key, lịch sử, kỹ năng) nằm trong thư mục `Wef` của Office: không xóa trong bất kỳ tệp cài, gỡ hay bước dọn nào.
