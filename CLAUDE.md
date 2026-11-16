# AI Đa Nguồn — add-in AI cho Word, Excel, PowerPoint

Add-in Office dùng API key của nhiều nhà cung cấp, AI tự thao tác trên tài liệu như add-in Claude chính thức. Dùng cá nhân (anh Dũng, pháp lý – bất động sản). Đang chạy **v4.2.1** (phần A của Đợt 3: kỹ năng và công cụ soạn văn bản hành chính đúng thể thức, cảnh báo khi AI nói đã sửa mà lượt đó không ghi gì, ưu tiên model miễn phí gọi công cụ đáng tin; phát hành 16/11/2026; phần B của Đợt 3 — ghi chéo ứng dụng và 7 công cụ PowerPoint — là bản **4.3.0**, mã và nghiệm thu đã xong trên nhánh `claude/dot-3-b-wip`, trạng thái phát hành xem đầu `tasks/HANDOFF.md`; bản 4.2.0 của Đợt 2: đăng nhập OpenRouter một lần bấm trong hộp thoại của Office, lối dán mã, thẻ hạn mức, Đăng xuất, thông báo key sai / hết hạn / hết tiền kèm nút Đổi nguồn; phát hành 16/11/2026; bản 4.1.0 của Đợt 1B: Word soát và tra — soát nhất quán bằng mã, đọc chú thích và đầu/chân trang, so sánh phiên bản, thể thức trang, PDF quét, Suy nghĩ kỹ; bản 4.0.0 của Đợt 1A: Word đọc đúng số điều, sửa đúng chỗ, có Hoàn tác). Kế hoạch v4 (bản 2) đã duyệt ngày 15/11/2026, đợt kế tiếp là 3; trước đó anh nêu hai việc (kỹ năng soạn thảo văn bản hành chính; model miễn phí soạn Word kém) — xem mục Bước tiếp theo của `tasks/HANDOFF.md`.

## Đọc trước khi làm bất cứ việc gì

1. `tasks/HANDOFF.md` — trạng thái, quyết định, vấn đề mở, bài học.
2. `tasks/ke-hoach-v4.md` — hiện trạng đã kiểm, phương án đăng nhập, kế hoạch theo đợt (0A, 0B, 1A, 1B, 2, 3, 4) và tiêu chí nghiệm thu. Bằng chứng rà soát: `tasks/ra-soat-2026-11-15/DOC-TRUOC.md`.
3. `tasks/dot-3.md` — Đợt 3 (đang làm): tiêu chí, bảng thăm dò Word / kênh lệnh chéo ứng dụng / PowerPoint, bảng nghiệm thu phần A; `tasks/dot-2.md` — Đợt 2 (đăng nhập và chi phí): bảng đối chiếu tài liệu chính thức, bảng nghiệm thu, và mục "Việc buộc phải có anh ở máy" ghi từng cú bấm; `tasks/dot-1b.md` — việc, bảng thăm dò Word và kết quả của Đợt 1B; `tasks/dot-1a.md` — Đợt 1A (Đợt 0A: `tasks/dot-0.md`, Đợt 0B: `tasks/dot-0b.md`, đã xong). Đợt kế tiếp là 3: viết `tasks/dot-3.md` trước khi làm.
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
  providers.ts models.ts storage.ts   nhà cung cấp (kèm cấu hình đăng nhập `oauth`), danh mục model, key lưu cục bộ
  oauth.ts                            đăng nhập để nhận key (OAuth PKCE), khối dùng chung: mã kiểm / thách thức, địa chỉ, đọc mã dán, đổi mã lấy key
  credits.ts                          hạn mức của key OpenRouter (GET /key, /credits): dòng hiển thị và mức cảnh báo
  errors.ts                           lỗi → câu tiếng Việt; phân loại lỗi key và tiền theo nhà cung cấp; maskSecrets che key trong mọi thông báo
  skills.ts skillpack.ts              kỹ năng (SKILL.md), 18 kỹ năng dựng sẵn
  adminDoc.ts                         bản vẽ văn bản hành chính đúng thể thức (công văn, tờ trình, quyết định…): bảng không viền + đoạn đã định dạng
  crossapp.ts                         ghi chéo ứng dụng: nhịp sống, hộp thư lệnh trong kho dùng chung, kiểm lệnh ở khung nhận (chỉ thêm mới)
  deckcheck.ts                        kiểm bài trình chiếu theo mẫu bằng mã: phông, cỡ chữ, tiêu đề, chữ tràn, hình ra ngoài, chồng nhau
  claims.ts                           nhận diện "AI nói đã sửa mà lượt đó không ghi gì" → cảnh báo trên khung (không vào lịch sử gửi model)
  outline.ts                          mục lục điều khoản: số do Word tự đánh + số gõ tay, tìm điều theo cách gọi ("khoản 2 Điều 5")
  worddiff.ts                         so khác mức từ, kế hoạch sửa một đoạn Word, chống sửa nhầm đoạn, Hoàn tác Word
  consistency.ts                      soát nhất quán bằng mã: đánh số, dẫn chiếu, thuật ngữ, tên các bên, chỗ trống, số lệch chữ
  compare.ts pageformat.ts            so hai phiên bản theo điều khoản; kiểm lề, khổ giấy, phông, cỡ chữ
  attachedFiles.ts                    bản gốc tệp .docx đính kèm trong phiên (cho công cụ so sánh phiên bản)
  conversations.ts db.ts share.ts attachments.ts refs.ts
src/office/          công cụ Office.js: excel*, word*, ppt*, common, selftest (nút Tự kiểm tra; Word ở selftestWord.ts, chạy trong vùng cô lập)
  crossExec.ts       khung nhận của ghi chéo: chạy thao tác ghi hộ bằng công cụ ghi của chính ứng dụng (có Hoàn tác)
  pptDesign.ts       7 công cụ PowerPoint của Đợt 3: ảnh slide, kiểm bài, đọc mẫu, định dạng theo đoạn, sửa bảng, vùng chọn, đổi bố cục
  wordAdmin.ts       2 công cụ Word của Đợt 3: dựng văn bản hành chính theo bản vẽ, đặt khổ giấy và lề
  wordReview.ts      4 công cụ Word chỉ đọc của Đợt 1B: soát nhất quán, đọc phần ngoài thân văn bản, thể thức trang, so sánh phiên bản
  dialog.ts          hộp thoại của Office (displayDialogAsync) và mở trình duyệt ngoài, dùng cho đăng nhập
src/dialog/          dang-nhap.html + dang-nhap.ts: trang mở trong hộp thoại của Office khi đăng nhập (đóng gói thành dist/dang-nhap.html)
src/taskpane/components/   App.tsx, Chat.tsx (điều phối + giao diện), Settings.tsx, SignInCard.tsx + useCredits.ts (thẻ OpenRouter: đăng nhập, dán mã, hạn mức, đăng xuất)
test/unit/           mocha, 420 kiểm thử (dữ liệu mẫu ở test/fixtures/)
cai-dat/             tệp .cmd cài và gỡ add-in vào Office (ghi registry WEF\Developer)
```

## Lệnh

```
npm ci                      # .npmrc phải trỏ registry.npmjs.org
npm test                    # 420/420 phải đạt
npx tsc --noEmit -p .       # không lỗi kiểu
npm run build               # ra dist/, manifest tự đổi sang địa chỉ GitHub Pages
```

Phát hành tự động: đẩy lên `main` là GitHub Actions (`.github/workflows/phat-hanh.yml`) kiểm thử, kiểm kiểu, đóng gói rồi đưa `dist/` lên `gh-pages`; trượt thì không phát hành. Nhánh `thu/**` chạy cùng quy trình nhưng không đẩy. Xem kết quả: https://github.com/DungHoang1985/ai-da-nguon/actions

Xem bản mới trong Office: đặt con trỏ trong khung add-in rồi Ctrl+F5 (bằng mã: `node kiem-thu/cdp.mjs reload 9229`). Số phiên bản hiện ở đầu khung.

Kiểm trong Office bằng mã (không cần điều khiển màn hình, dùng được cả khi anh không ngồi trước máy): `kiem-thu/cdp.mjs`, cách dùng ở đầu tệp và ở mục Bài học của `tasks/HANDOFF.md`; các bài thử mẫu ở `kiem-thu/cdp-bai/`. Mở và đóng Office thử: `kiem-thu/mo-office.ps1` (tự dừng nếu Office đang mở tài liệu không phải tệp thử). **Thử mã mới trước khi phát hành**: `npm run build` rồi `node kiem-thu/cdp.mjs serve 9229 dist` (chạy nền) nạp bản cục bộ vào khung add-in, không đẩy gì lên `main`. Tệp Word thử: `kiem-thu/docx-thu.mjs` (thăm dò đánh số), `kiem-thu/hop-dong-thu.mjs` (hợp đồng 50 trang có sửa đổi và nhận xét của "Đối tác"), `kiem-thu/soat-thu.mjs` (hợp đồng sạch / cài sẵn 9 lỗi / phiên bản khác lệch 10 chỗ / công văn lề sai, có chú thích, đầu và chân trang, bookmark), `kiem-thu/pdf-quet-thu.mjs` (PDF quét 3 trang). Xem bản in của văn bản Word mà không cần điều khiển màn hình: Word COM `ExportAsFixedFormat` ra PDF rồi `pypdfium2` ra PNG (cách làm ở mục Bài học của HANDOFF). Bài thử Đợt 3: `kiem-thu/cdp-bai/dot-3-*.js`, `tham-do-3-*.js`; bài PowerPoint thử `kiem-thu/ppt-thu.mjs` (10 slide theo mẫu riêng, cài 5 lỗi lệch mẫu và 1 slide chữ tràn). Mở ba ứng dụng cùng lúc: mở Word bằng `mo-office.ps1`, rồi `Start-Process excel.exe` / `powerpnt.exe` kèm biến môi trường cổng gỡ lỗi; mỗi khung một tiến trình `cdp.mjs serve 9229/<Ứng dụng> dist`. Hộp thoại đăng nhập là một trang riêng trên cùng cổng: `node kiem-thu/cdp.mjs list 9229` liệt kê các trang (giá trị tham số nhạy cảm chỉ in độ dài), `9229#openrouter.ai` hoặc `9229#dang-nhap` chọn trang đó cho `eval`, `shot`, `nav`; `serve` tự gắn vào hộp thoại để nó cũng nhận bản cục bộ. Bài thử Đợt 2: `kiem-thu/cdp-bai/dot-2-*.js`.

## Quy tắc làm việc với anh Dũng

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
