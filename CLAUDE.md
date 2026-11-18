# AI Đa Nguồn — add-in AI cho Word, Excel, PowerPoint

Add-in Office dùng API key của nhiều nhà cung cấp, AI tự thao tác trên tài liệu như add-in Claude chính thức. Dùng cá nhân (anh Dũng, pháp lý – bất động sản). Đang chạy **v4.7.1** (bản vá 18/11/2026: chữ có “ư” và tiêu đề slide biểu đồ trong PowerPoint không còn thành Times New Roman; xem mục "Sau phát hành 4.7.0" của `tasks/dot-7.md`). Bản trước: v4.7.0 (Đợt 7, phát hành 18/11/2026: nút Làm tiếp khi chạm trần số bước, thêm nhiều slide một lệnh, đặt ảnh đính kèm vào slide, nhật ký Excel ghi giá trị cũ / mới, chế độ thử không lưu lịch sử, sửa theo ba lượt soát mã độc lập; xem `tasks/dot-7.md`). Bản trước: v4.6.0 (Đợt 6, phát hành 17/11/2026: thêm nguồn nói cùng kiểu với Anthropic — OpenRouter kênh Claude và DeepSeek kiểu Anthropic dùng chung key đã có, MiniMax, Kimi, và mục "nguồn tự khai (kiểu Anthropic)" trong Cài đặt; chi tiết và phần chưa chạy thật ở `tasks/dot-6.md`). Bản trước: 4.5.0 (Đợt 5, phát hành 17/11/2026: AI đọc tệp PDF đính kèm theo trang — tệp quét dài và tệp lẫn trang quét; đoạn Word có sửa đổi chờ duyệt được đưa cả chữ gốc và chữ đang sửa; lý do kèm mỗi chỗ sửa Word; dòng nhắc riêng tư khi tệp đi qua OpenRouter hoặc gói miễn phí; chi tiết ở `tasks/dot-5.md`). Bản trước: 4.4.1 (bản vá 17/11/2026: Gemini hết bị từ chối trong Word — lược đồ công cụ chỉ còn enum chữ; khung hết treo khi nhà cung cấp im lặng — báo lỗi sau 150 giây; kết quả so song song với add-in Claude chính thức ở `tasks/so-song-song.md`). Bản trước: 4.4.0 (Đợt 4, phát hành 16/11/2026: Excel và dấu vết soát lại — sheet "Nhật ký AI" và tô sáng ô vừa sửa, mặc định tắt; sửa biểu đồ và PivotTable có sẵn; chuẩn bị in; kiểm toán mô hình; giải ngược và bảng nhạy cảm; làm sạch dữ liệu; khóa ô công thức; hỏi trước công thức nguy hiểm). Bản trước: 4.3.1 (bản vá 16/11/2026: một nhiệm vụ ít lượt gọi AI hơn — gộp lệnh, chụp nhiều slide một lần, khung hiện số lần gọi; bản 4.3.0 là Đợt 3, phát hành cùng ngày: ghi chéo ứng dụng Excel – Word – PowerPoint luôn hỏi trước, chỉ báo kết nối, 7 công cụ PowerPoint mới, AI tự nhìn ảnh slide; việc kế tiếp anh nêu: giảm số lần gọi AI cho một nhiệm vụ — xem `tasks/HANDOFF.md`; bản 4.2.1 là phần A của Đợt 3: kỹ năng và công cụ soạn văn bản hành chính đúng thể thức, cảnh báo khi AI nói đã sửa mà lượt đó không ghi gì, ưu tiên model miễn phí gọi công cụ đáng tin; phát hành 16/11/2026; phần B của Đợt 3 — ghi chéo ứng dụng và 7 công cụ PowerPoint — là bản **4.3.0**, mã và nghiệm thu đã xong trên nhánh `claude/dot-3-b-wip`, trạng thái phát hành xem đầu `tasks/HANDOFF.md`; bản 4.2.0 của Đợt 2: đăng nhập OpenRouter một lần bấm trong hộp thoại của Office, lối dán mã, thẻ hạn mức, Đăng xuất, thông báo key sai / hết hạn / hết tiền kèm nút Đổi nguồn; phát hành 16/11/2026; bản 4.1.0 của Đợt 1B: Word soát và tra — soát nhất quán bằng mã, đọc chú thích và đầu/chân trang, so sánh phiên bản, thể thức trang, PDF quét, Suy nghĩ kỹ; bản 4.0.0 của Đợt 1A: Word đọc đúng số điều, sửa đúng chỗ, có Hoàn tác). Kế hoạch v4 (bản 2) đã duyệt ngày 15/11/2026; các đợt 0A → 4 đã làm xong, còn Đợt 5 là tùy chọn (làm khi anh cần) và các việc nợ — xem mục Bước tiếp theo của `tasks/HANDOFF.md`.

## Đọc trước khi làm bất cứ việc gì

1. `tasks/HANDOFF.md` — trạng thái, quyết định, vấn đề mở, bài học.
2. `tasks/ke-hoach-v4.md` — hiện trạng đã kiểm, phương án đăng nhập, kế hoạch theo đợt (0A, 0B, 1A, 1B, 2, 3, 4) và tiêu chí nghiệm thu. Bằng chứng rà soát: `tasks/ra-soat-2026-11-15/DOC-TRUOC.md`.
3. `tasks/dot-6.md` — Đợt 6 (nguồn nói kiểu Anthropic: bảng đo từng nguồn có cho add-in gọi không, thiết kế `dialect`, bảng nghiệm thu, việc cần anh: key và tiền ở từng nguồn); `tasks/dot-5.md` — Đợt 5 (anh chọn 17/11/2026: đọc tài liệu quét theo trang, kiểm nốt, hai việc nhỏ từ buổi so song song; đã phát hành trong bản 4.5.0; phần D — các việc tùy chọn còn lại — chờ anh chọn); `tasks/dot-4.md` — Đợt 4 (Excel và dấu vết soát lại): tiêu chí, bảng thăm dò Excel, bảng nghiệm thu, bảng soát mã (đã sửa / để lại); `tasks/dot-3.md` — Đợt 3: tiêu chí, bảng thăm dò Word / kênh lệnh chéo ứng dụng / PowerPoint, bảng nghiệm thu phần A; `tasks/dot-2.md` — Đợt 2 (đăng nhập và chi phí): bảng đối chiếu tài liệu chính thức, bảng nghiệm thu, và mục "Việc buộc phải có anh ở máy" ghi từng cú bấm; `tasks/dot-1b.md` — việc, bảng thăm dò Word và kết quả của Đợt 1B; `tasks/dot-1a.md` — Đợt 1A (Đợt 0A: `tasks/dot-0.md`, Đợt 0B: `tasks/dot-0b.md`, đã xong). Đợt kế tiếp là 5 (tùy chọn): hỏi anh chọn việc, viết `tasks/dot-5.md` trước khi làm.
4. `docs/so-sanh-claude-chinh-thuc.md` — bảng so sánh 67 khả năng với bản chính thức.

## Nguồn mã

Từ 15/11/2026 (bước 1 của `tasks/dot-0.md`), thư mục này là bản làm việc của nhánh `main` — một nguồn duy nhất, gồm cả `docs/`, `tasks/`, `kiem-thu/` và tệp này.

- Kho: https://github.com/DungHoang1985/ai-da-nguon (công khai) — nhánh `main` (mã), nhánh `gh-pages` (bản đang chạy).
- Bản chạy: https://dunghoang1985.github.io/ai-da-nguon/
- Trước khi sửa: `git status` sạch và `git pull` để lấy bản mới nhất.

## Cấu trúc mã (nhánh `main`)

```
src/core/            lõi không phụ thuộc Office, kiểm thử được bằng npm test
  agent/turn.ts      điều phối một lượt (tóm tắt → dựng lời nhắc → chạy tác tử hoặc trò chuyện); giao diện chỉ gọi runTurn
  agent/loop.ts      vòng lặp tác tử (gọi model → chạy công cụ → lặp); lịch sử chỉ thêm, hạn chờ công cụ
  agent/llm.ts       gọi model, gom 3 kiểu API: OpenAI, Anthropic, Gemini; tự thử lại, bộ nhớ đệm Anthropic, đếm token; gửi PDF quét nguyên tệp; giữ và gửi lại khối suy nghĩ
  agent/thinking.ts  công tắc "Suy nghĩ kỹ": tham số suy nghĩ theo từng nguồn và đời model; nguồn nào đọc được PDF quét
  agent/prompt.ts    lời nhắc hệ thống cố định + quy tắc an toàn; trạng thái tài liệu đi trong tin nhắn của lượt
  agent/compaction.ts  tự tóm tắt hội thoại dài, ngưỡng theo sức chứa model
  providers.ts models.ts storage.ts   nhà cung cấp (kèm cấu hình đăng nhập `oauth`; `dialect` cho nguồn nói kiểu Anthropic, `keyFrom` dùng chung key, nguồn tự khai — Đợt 6), danh mục model, key lưu cục bộ
  oauth.ts                            đăng nhập để nhận key (OAuth PKCE), khối dùng chung: mã kiểm / thách thức, địa chỉ, đọc mã dán, đổi mã lấy key
  credits.ts                          hạn mức của key OpenRouter (GET /key, /credits): dòng hiển thị và mức cảnh báo
  errors.ts                           lỗi → câu tiếng Việt; phân loại lỗi key và tiền theo nhà cung cấp; maskSecrets che key trong mọi thông báo
  skills.ts skillpack.ts              kỹ năng (SKILL.md), 18 kỹ năng dựng sẵn
  adminDoc.ts                         bản vẽ văn bản hành chính đúng thể thức (công văn, tờ trình, quyết định…): bảng không viền + đoạn đã định dạng
  crossapp.ts                         ghi chéo ứng dụng: nhịp sống, hộp thư lệnh trong kho dùng chung, kiểm lệnh ở khung nhận (chỉ thêm mới)
  deckcheck.ts                        kiểm bài trình chiếu theo mẫu bằng mã: phông, cỡ chữ, tiêu đề, chữ tràn, hình ra ngoài, chồng nhau
  auditlog.ts                         dấu vết soát lại của Excel: dòng của sheet "Nhật ký AI", chữ ghi vào nhật ký luôn là chữ, nhận diện sheet nhật ký
  formularisk.ts                      công thức Excel phải hỏi trước khi ghi: WEBSERVICE, INDIRECT, DDE, liên kết tệp ngoài…
  modelaudit.ts goalseek.ts           kiểm toán mô hình trên lưới R1C1 (số gõ cứng, công thức lệch, tỷ lệ gõ trong công thức); giải ngược tìm đầu vào
  claims.ts                           nhận diện "AI nói đã sửa mà lượt đó không ghi gì" → cảnh báo trên khung (không vào lịch sử gửi model)
  outline.ts                          mục lục điều khoản: số do Word tự đánh + số gõ tay, tìm điều theo cách gọi ("khoản 2 Điều 5")
  worddiff.ts                         so khác mức từ, kế hoạch sửa một đoạn Word, chống sửa nhầm đoạn, Hoàn tác Word
  consistency.ts                      soát nhất quán bằng mã: đánh số, dẫn chiếu, thuật ngữ, tên các bên, chỗ trống, số lệch chữ
  compare.ts pageformat.ts            so hai phiên bản theo điều khoản; kiểm lề, khổ giấy, phông, cỡ chữ
  attachedFiles.ts                    bản gốc tệp .docx và .pdf đính kèm trong phiên (cho công cụ so sánh phiên bản và đọc theo trang)
  pdfPages.ts pdfRender.ts            đọc PDF đính kèm theo trang (Đợt 5): chọn trang, trang quét / trang chữ, đường đi của PDF quét, lời dặn chép nguyên văn; dựng ảnh trang bằng pdf.js
  conversations.ts db.ts share.ts attachments.ts refs.ts
src/office/          công cụ Office.js: excel*, word*, ppt*, common, selftest (nút Tự kiểm tra; Word ở selftestWord.ts, chạy trong vùng cô lập)
  excelModel.ts      9 công cụ Excel của Đợt 4: liệt kê và sửa biểu đồ / Pivot có sẵn, chuẩn bị in, tên vùng, kiểm toán mô hình, giải ngược, làm sạch, bảo vệ
  excelTrail.ts      lớp bọc bộ công cụ Excel: nhật ký, tô sáng, chặn AI ghi vào sheet nhật ký (getHostTools luôn trả bộ đã bọc)
  crossExec.ts       khung nhận của ghi chéo: chạy thao tác ghi hộ bằng công cụ ghi của chính ứng dụng (có Hoàn tác)
  pptDesign.ts       7 công cụ PowerPoint của Đợt 3: ảnh slide, kiểm bài, đọc mẫu, định dạng theo đoạn, sửa bảng, vùng chọn, đổi bố cục
  wordAdmin.ts       2 công cụ Word của Đợt 3: dựng văn bản hành chính theo bản vẽ, đặt khổ giấy và lề
  wordReview.ts      4 công cụ Word chỉ đọc của Đợt 1B: soát nhất quán, đọc phần ngoài thân văn bản, thể thức trang, so sánh phiên bản
  dialog.ts          hộp thoại của Office (displayDialogAsync) và mở trình duyệt ngoài, dùng cho đăng nhập
src/dialog/          dang-nhap.html + dang-nhap.ts: trang mở trong hộp thoại của Office khi đăng nhập (đóng gói thành dist/dang-nhap.html)
src/taskpane/components/   App.tsx, Chat.tsx (điều phối + giao diện), Settings.tsx, SignInCard.tsx + useCredits.ts (thẻ OpenRouter: đăng nhập, dán mã, hạn mức, đăng xuất)
test/unit/           mocha, 524 kiểm thử (dữ liệu mẫu ở test/fixtures/)
cai-dat/             tệp .cmd cài và gỡ add-in vào Office (ghi registry WEF\Developer)
```

## Lệnh

```
npm ci                      # .npmrc phải trỏ registry.npmjs.org
npm test                    # 524/524 phải đạt
npx tsc --noEmit -p .       # không lỗi kiểu
npm run build               # ra dist/, manifest tự đổi sang địa chỉ GitHub Pages
```

Phát hành tự động: đẩy lên `main` là GitHub Actions (`.github/workflows/phat-hanh.yml`) kiểm thử, kiểm kiểu, đóng gói rồi đưa `dist/` lên `gh-pages`; trượt thì không phát hành. Nhánh `thu/**` chạy cùng quy trình nhưng không đẩy. Xem kết quả: https://github.com/DungHoang1985/ai-da-nguon/actions

Xem bản mới trong Office: đặt con trỏ trong khung add-in rồi Ctrl+F5 (bằng mã: `node kiem-thu/cdp.mjs reload 9229`). Số phiên bản hiện ở đầu khung.

Kiểm trong Office bằng mã (không cần điều khiển màn hình, dùng được cả khi anh không ngồi trước máy): `kiem-thu/cdp.mjs`, cách dùng ở đầu tệp và ở mục Bài học của `tasks/HANDOFF.md`; các bài thử mẫu ở `kiem-thu/cdp-bai/`. Mở và đóng Office thử: `kiem-thu/mo-office.ps1` — **mặc định mở trên desktop ẩn** (`kiem-thu/desktop-rieng.ps1`), không hiện lên màn hình của anh; tự dừng nếu anh đang mở Office ở desktop chính; `-HienLen` chỉ dùng khi anh đã đồng ý. Lệnh COM tới Office thử phải chạy bên trong desktop ẩn (`desktop-rieng.ps1 -Lenh powershell.exe -ThamSo …`); giới hạn và bảng thăm dò ở `tasks/HANDOFF.md`. **Thử mã mới trước khi phát hành**: `npm run build` rồi `node kiem-thu/cdp.mjs serve 9229 dist` (chạy nền) nạp bản cục bộ vào khung add-in, không đẩy gì lên `main`. Tệp Word thử: `kiem-thu/docx-thu.mjs` (thăm dò đánh số), `kiem-thu/hop-dong-thu.mjs` (hợp đồng 50 trang có sửa đổi và nhận xét của "Đối tác"), `kiem-thu/soat-thu.mjs` (hợp đồng sạch / cài sẵn 9 lỗi / phiên bản khác lệch 10 chỗ / công văn lề sai, có chú thích, đầu và chân trang, bookmark), `kiem-thu/pdf-quet-thu.mjs` (PDF quét 3 trang). Xem bản in của văn bản Word mà không cần điều khiển màn hình: Word COM `ExportAsFixedFormat` ra PDF rồi `pypdfium2` ra PNG (cách làm ở mục Bài học của HANDOFF). Bài thử Đợt 3: `kiem-thu/cdp-bai/dot-3-*.js`, `tham-do-3-*.js`; bài PowerPoint thử `kiem-thu/ppt-thu.mjs` (10 slide theo mẫu riêng, cài 5 lỗi lệch mẫu và 1 slide chữ tràn). Mở ba ứng dụng cùng lúc: mở Word bằng `mo-office.ps1`, rồi `Start-Process excel.exe` / `powerpnt.exe` kèm biến môi trường cổng gỡ lỗi; mỗi khung một tiến trình `cdp.mjs serve 9229/<Ứng dụng> dist`. Hộp thoại đăng nhập là một trang riêng trên cùng cổng: `node kiem-thu/cdp.mjs list 9229` liệt kê các trang (giá trị tham số nhạy cảm chỉ in độ dài), `9229#openrouter.ai` hoặc `9229#dang-nhap` chọn trang đó cho `eval`, `shot`, `nav`; `serve` tự gắn vào hộp thoại để nó cũng nhận bản cục bộ. Bài thử Đợt 2: `kiem-thu/cdp-bai/dot-2-*.js`. Bài thử Đợt 4 (Excel): `dot-4-excel-du-lieu.js` dựng ba sheet thử (bảng tiến độ, mô hình cài 4 lỗi, bảng hồ sơ có dòng trùng), `dot-4-excel-bai.js` giao việc rồi đọc lại sổ tính và số lần gọi AI, `dot-4-cong-tac.js` bật / tắt hai công tắc nhật ký, `dot-4-word-gui.js` (Word thu nhỏ, Groq 413), `tham-do-4-excel.js`.

## Quy tắc làm việc với anh Dũng

- **Không chiếm màn hình của anh** (anh đặt 17/11/2026, áp dụng mọi dự án — bản đầy đủ ở chỉ dẫn chung của máy): kiểm thử hay thao tác phải bật cửa sổ, bật Office thì chạy không cửa sổ hoặc trên desktop riêng; không bật gì lên desktop anh đang làm, không dùng công cụ điều khiển màn hình khi chưa hỏi; không có cách tách thì dừng và hỏi.
- Trả lời bằng tiếng Việt, kết luận trước, giải thích ngắn. Không kể lại các bước đã làm.
- **Anh không gõ lệnh, không cài tay.** Mọi việc Claude tự làm và tự kiểm. Việc buộc phải có anh (đăng nhập GitHub, bấm xác nhận) thì gom lại, nói rõ từng cú bấm.
- Mọi câu hỏi cần anh trả lời hoặc duyệt: dùng ô tích chọn, 2–4 phương án, phương án khuyến nghị đặt đầu và ghi "(Khuyến nghị)", ngôn ngữ dễ hiểu, tối đa 4 câu một lần. Điều tự kiểm được thì không hỏi.
- Trước khi làm: nêu tiêu chí nghiệm thu. Khi giao: bảng Tiêu chí · Kết quả (Đạt / Không đạt / Chưa kiểm được) · Bằng chứng. Không có bằng chứng thì không ghi Đạt.
- Bước không đạt: chỉ làm lại bước đó, tối đa 3 lần, rồi dừng và báo.
- Cuối mỗi phiên: cập nhật `tasks/HANDOFF.md` (đã xong · đang làm · quyết định và lý do · vấn đề mở · bước tiếp theo · bài học) và đề xuất bước tiếp theo.
- Anh sửa kết quả hoặc lỗi lặp lại: ghi vào mục Bài học của HANDOFF.

## Quy tắc kỹ thuật

- "Đạt" nghĩa là AI **thao tác thật** trên tài liệu trong Office. Ba mức bằng chứng, ghi đúng tên: **Tự kiểm tra** (nút trong add-in, chạy trong Office) · **Chạy thật** (giao việc cho model thật) · **Kiểm thử máy** (`npm test`, ngoài Office).
- Thêm hoặc sửa công cụ Office: phải thêm bước tương ứng vào `src/office/selftest.ts`.
- Word: mọi sửa đổi ghi bằng Track Changes. Thao tác rủi ro (ghi đè ô có dữ liệu, xóa ô/đoạn/slide/sheet, chấp nhận hoặc từ chối tất cả sửa đổi) luôn hỏi người dùng kèm lý do.
- Word, đã thăm dò trong Office (bảng ở `tasks/dot-1a.md`), không làm trái: `paragraph.text` và `search()` gồm cả chữ đã xóa có theo dõi → đọc bản hiện tại bằng `getText()`, định vị trên chuỗi `.text`; không thay cả đoạn trên đoạn đang có sửa đổi chờ duyệt; công cụ sửa theo đoạn phải đối chiếu `expect`; công cụ ghi mới phải ghi được Hoàn tác (xem `execute` trong `wordTools.ts`); bước Tự kiểm tra Word chỉ đụng vùng thử.
- Word, đã thăm dò ở Đợt 1B (bảng ở `tasks/dot-1b.md`), không làm trái: `paragraph.text` có ký tự U+0002 tại mốc chú thích (chữ đưa cho AI phải bỏ mốc; khi sửa đoạn coi mốc là chữ "không sống", đoạn có chú thích không thay cả đoạn); so sánh phiên bản không được ghi vào tài liệu đang mở (`compareFromBase64` với đích là tài liệu đang mở làm mất sửa đổi chờ duyệt) — đọc tệp kia bằng `application.createDocument`, redline chỉ mở ở tài liệu mới.
- Word, đã thăm dò ở Đợt 3 (bảng ở `tasks/dot-3.md`), không làm trái: thêm đoạn vào ô bảng phải chèn `"Before"` đoạn sẵn có (`"After"` rơi sang ô kế); đường kẻ = viền dưới `paragraph.borders.items[2]` của một đoạn rỗng và không bao giờ là đoạn cuối của ô (viền dính sang ô cùng hàng); không chèn bảng ngay sau một đoạn rỗng (bảng nuốt đoạn đó); Hoàn tác của phần chèn có bảng là gỡ thẳng theo định danh từ cuối lên; chèn có theo dõi nhưng định dạng phần mới chèn khi đã tắt theo dõi (tránh mỗi đoạn một sửa đổi "Formatted").
- Cảnh báo "AI nói đã sửa mà không ghi" (`claims.ts`) chỉ là phần hiển thị: không ghi vào lịch sử gửi model, không đổi lời nhắc hệ thống; ưu tiên không báo nhầm — thêm mẫu câu phải kèm kiểm thử cho cả ca phải bắt và ca dễ báo nhầm.
- Ghi chéo ứng dụng (Đợt 3), không làm trái: `write_other_app` luôn hỏi người dùng ở khung gửi (cờ `alwaysAsk`; "Cho phép hết lần này" và việc tắt "hỏi trước thao tác rủi ro" đều không bỏ qua được) và chỉ gửi đúng nội dung vừa hỏi; hộp hỏi phải hiện gần đủ nội dung và tên tệp đích; khung nhận chỉ chạy các thao tác THÊM mới trong danh sách đóng của `crossapp.ts`, kiểm lại lệnh, ghi bằng công cụ của chính nó (Word: Track Changes) và để lại mục có nút Hoàn tác; nội dung lệnh không bao giờ đưa cho model ở khung nhận; công thức từ ứng dụng khác ghi thành chữ. Thêm thao tác ghi chéo mới phải qua soát mã độc lập.
- Ảnh do công cụ chụp (`__images` trong kết quả công cụ) được vòng lặp tách ra, gửi ở một tin nhắn user riêng có cờ `fromTool` và câu "không phải yêu cầu mới của người dùng"; không nhét ảnh vào kết quả dạng chữ, không lưu ảnh xuống kho.
- PowerPoint, đã thăm dò ở Đợt 3 (bảng ở `tasks/dot-3.md`): không có API cho ghi chú diễn giả, số liệu của biểu đồ đã có, phông chủ đề (đọc từ gói của slide xuất ra), xóa bố cục; hình trùng tên phải chỉ bằng số thứ tự; số đo đưa cho AI phải sát với cái mắt thấy (so vùng chữ, không so khung rỗng).
- Excel, Đợt 4, không làm trái (bảng thăm dò ở `tasks/dot-4.md`): nhật ký và tô sáng **mặc định tắt**, lúc bật phải có câu cảnh báo sheet đi theo tệp; mọi chữ ghi vào sheet "Nhật ký AI" phải qua `logText` (không bao giờ thành công thức); AI không ghi, sửa, đổi tên sheet nhật ký (lớp chặn so tên SAU khi làm sạch và tính cả lệnh không nêu sheet), chỉ được xóa cả sheet khi người dùng nhờ; nhật ký không vào lịch sử gửi model; bộ công cụ Excel chỉ dùng qua `getHostTools` (đã bọc `withAuditTrail`, gỡ `__trail` khỏi kết quả); `excel_what_if` luôn trả ô đầu vào về nội dung gốc và không tự ghi kết quả; `excel_protect` không bao giờ đặt mật khẩu; phép kiểm rủi ro của công cụ ghi phải chạy trên dữ liệu đã chuẩn hóa bằng đúng hàm công cụ dùng (`normalize2D`); Tự kiểm tra Excel chạy với nhật ký tắt và sheet nhật ký thử riêng.
- Tệp đính kèm theo trang (Đợt 5, `tasks/dot-5.md`), không làm trái: ảnh trang quét chỉ gửi qua kênh ảnh công cụ (`__images`), tối đa 4 trang một lần gọi; lời dặn chép nguyên văn nằm trong kết quả công cụ, không vào lời nhắc hệ thống; nguồn không nhận PDF mà danh mục không ghi rõ model nhận ảnh thì KHÔNG gửi ảnh trang (`canSeePages` — gửi là bị từ chối cả lượt); PDF quét từ 5 trang trở xuống với nguồn đọc được nguyên tệp vẫn gửi nguyên tệp; `read_attachment_pages` luôn có trong bộ công cụ chung (không thêm bớt theo lượt).
- Word, Đợt 5: kết quả đọc và tìm kiếm đưa chữ gốc (`before`) và người sửa (`changed_by`) cho đoạn có sửa đổi chờ duyệt kèm `pending_note`; tham số `reason` của công cụ sửa Word chỉ để hiển thị (nhãn bước, hộp hỏi, mục Hoàn tác), không ghi vào tài liệu.
- Nguồn nói kiểu Anthropic (Đợt 6, `tasks/dot-6.md`), không làm trái: yêu cầu gửi `api.anthropic.com` không được đổi (có kiểm thử chặn); khác biệt của từng nguồn chỉ khai trong `dialect`, không rẽ nhánh theo mã nguồn trong `llm.ts`; nguồn `bare` chỉ gửi tiêu đề key và `content-type` (thêm `anthropic-version` là bị trình duyệt chặn ở MiniMax, Kimi, OpenRouter); nguồn `keyFrom` không lưu bản sao key; nguồn tự khai chỉ nhận `https`, mã phải có tiền tố `tuy-chinh-`, không đè được nguồn có sẵn, và thẻ phải hiện tên miền nhận key; thêm sẵn một nguồn mới chỉ sau khi đọc tài liệu chính thức và đo bằng trình duyệt thật (Edge `--headless=new`, cả bước dò trước với đúng bộ tiêu đề).
- Kiểm toán mô hình (`modelaudit.ts`) và công thức cần hỏi (`formularisk.ts`) ưu tiên không báo nhầm: thêm mẫu phải kèm kiểm thử cho cả ca phải bắt và ca dễ báo nhầm (tham chiếu bảng `[2024]`, hàng số nhập tay xen công thức, gạch đầu dòng "- …").
- `maskSecrets` che cả mã tổ chức / dự án của tài khoản (`org_…`, `proj_…`): bằng chứng và câu lỗi không được lộ.
- Model mặc định và dấu ⚠ "đã thấy không gọi công cụ" (`catalog.ts`) chỉ đổi khi có bằng chứng Chạy thật và anh đã chọn.
- Khối suy nghĩ của model (Anthropic `thinking`, OpenRouter `reasoning_details`) nằm trong `raw` của tin nhắn trợ lý và được gửi lại nguyên văn, đúng thứ tự; chỗ nào buộc phải đổi tin nhắn cũ (tóm tắt, rút gọn khi lưu) thì bỏ khối suy nghĩ đi kèm. Tham số suy nghĩ không được nằm trong lời nhắc hệ thống, bộ công cụ hay tin nhắn.
- Kỹ năng dựng sẵn: không ghi số hiệu văn bản pháp luật (có kiểm thử chặn).
- Lịch sử hội thoại gửi model là **chỉ thêm**: không sửa, không cắt tin nhắn cũ giữa hai lần tóm tắt; lời nhắc hệ thống không chứa thứ hay đổi (ngày giờ, trạng thái tài liệu). Sửa khác đi là mất bộ nhớ đệm lời nhắc (tốn tiền) — có kiểm thử chặn trong `test/unit/dot0b.test.ts`.
- Tài liệu của nhà cung cấp và của Office.js thay đổi nhanh: đối chiếu tài liệu chính thức hiện hành trước khi viết mã gọi API, không viết theo trí nhớ.
- Không thêm máy chủ riêng, không đổi khung giao diện, không viết lại từ đầu (đã cân nhắc và loại ở mục A của kế hoạch).
- Kho GitHub là **công khai**: không đưa key, token, thông tin riêng tư vào bất kỳ tệp nào.
- Đăng nhập và khóa (Đợt 2), không làm trái: mã kiểm của luồng đăng nhập (`code_verifier`) chỉ nằm trong bộ nhớ của khung, không ghi kho lưu, không vào địa chỉ trang; trang hộp thoại `dang-nhap.html` phải cùng tên miền với khung và chỉ chuyển tới tên miền khai trong cấu hình `oauth` (không nhận địa chỉ đích từ tham số); khung chỉ nhận mã khi `state` khớp lần đăng nhập đang mở; mọi câu lỗi của nhà cung cấp đưa ra ngoài (hiện, chép, lưu cùng hội thoại) phải qua `maskSecrets`; key OpenRouter kiểm bằng `GET /key` trước khi lưu (`/models` trả 200 cả khi key sai); lỗi đã rõ lý do về key hoặc tiền (`ApiError.reason`) không tự thử lại. Bài thử không đọc kho lưu của khung, không bấm nút mở trình duyệt ngoài khi chưa báo anh, và che số trong thẻ hạn mức trước khi ghi bằng chứng.
- Tài khoản của anh: phiên không đăng nhập thay anh, không tạo hay thu hồi key, không nhận tín dụng, không bấm đồng ý điều khoản, không xóa key thật đang lưu khi chưa hỏi. Việc đó gom vào mục "Việc buộc phải có anh ở máy" của tệp đợt, ghi rõ từng cú bấm.
- Các bài học Office.js đã trả giá nằm ở cuối `tasks/HANDOFF.md`; đọc trước khi đụng tới Word, PowerPoint hoặc WebView2.
- Dữ liệu add-in (key, lịch sử, kỹ năng) nằm trong thư mục `Wef` của Office, tách theo tài khoản Office: không xóa thư mục này trong bất kỳ tệp cài, gỡ hay bước dọn nào.
