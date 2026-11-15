# AI Đa Nguồn — add-in AI cho Word, Excel, PowerPoint

Add-in Office dùng API key của nhiều nhà cung cấp, AI tự thao tác trên tài liệu như add-in Claude chính thức. Dùng cá nhân (anh Dũng, pháp lý – bất động sản). Đang chạy **v4.0.0** (Đợt 1A: Word đọc đúng số điều, sửa đúng chỗ, có Hoàn tác; phát hành 15/11/2026). **Bản 4.1.0** (Đợt 1B: Word soát và tra — soát nhất quán bằng mã, đọc chú thích và đầu/chân trang, so sánh phiên bản, thể thức trang, PDF quét, Suy nghĩ kỹ) đã xong mã và kiểm trên bản cục bộ, **chờ anh duyệt phát hành** (xem `tasks/HANDOFF.md`). Kế hoạch v4 (bản 2) đã duyệt ngày 15/11/2026, đợt kế tiếp là 2.

## Đọc trước khi làm bất cứ việc gì

1. `tasks/HANDOFF.md` — trạng thái, quyết định, vấn đề mở, bài học.
2. `tasks/ke-hoach-v4.md` — hiện trạng đã kiểm, phương án đăng nhập, kế hoạch theo đợt (0A, 0B, 1A, 1B, 2, 3, 4) và tiêu chí nghiệm thu. Bằng chứng rà soát: `tasks/ra-soat-2026-11-15/DOC-TRUOC.md`.
3. `tasks/dot-1b.md` — việc, bảng thăm dò Word và kết quả của Đợt 1B; `tasks/dot-1a.md` — Đợt 1A (Đợt 0A: `tasks/dot-0.md`, Đợt 0B: `tasks/dot-0b.md`, đã xong). Đợt kế tiếp là 2: viết `tasks/dot-2.md` trước khi làm.
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
  providers.ts models.ts storage.ts   nhà cung cấp, danh mục model, key lưu cục bộ
  skills.ts skillpack.ts              kỹ năng (SKILL.md), 17 kỹ năng dựng sẵn
  outline.ts                          mục lục điều khoản: số do Word tự đánh + số gõ tay, tìm điều theo cách gọi ("khoản 2 Điều 5")
  worddiff.ts                         so khác mức từ, kế hoạch sửa một đoạn Word, chống sửa nhầm đoạn, Hoàn tác Word
  consistency.ts                      soát nhất quán bằng mã: đánh số, dẫn chiếu, thuật ngữ, tên các bên, chỗ trống, số lệch chữ
  compare.ts pageformat.ts            so hai phiên bản theo điều khoản; kiểm lề, khổ giấy, phông, cỡ chữ
  attachedFiles.ts                    bản gốc tệp .docx đính kèm trong phiên (cho công cụ so sánh phiên bản)
  conversations.ts db.ts share.ts attachments.ts refs.ts
src/office/          công cụ Office.js: excel*, word*, ppt*, common, selftest (nút Tự kiểm tra; Word ở selftestWord.ts, chạy trong vùng cô lập)
  wordReview.ts      4 công cụ Word chỉ đọc của Đợt 1B: soát nhất quán, đọc phần ngoài thân văn bản, thể thức trang, so sánh phiên bản
src/taskpane/components/   App.tsx, Chat.tsx (điều phối + giao diện), Settings.tsx
test/unit/           mocha, 291 kiểm thử (dữ liệu mẫu ở test/fixtures/)
cai-dat/             tệp .cmd cài và gỡ add-in vào Office (ghi registry WEF\Developer)
```

## Lệnh

```
npm ci                      # .npmrc phải trỏ registry.npmjs.org
npm test                    # 291/291 phải đạt
npx tsc --noEmit -p .       # không lỗi kiểu
npm run build               # ra dist/, manifest tự đổi sang địa chỉ GitHub Pages
```

Phát hành tự động: đẩy lên `main` là GitHub Actions (`.github/workflows/phat-hanh.yml`) kiểm thử, kiểm kiểu, đóng gói rồi đưa `dist/` lên `gh-pages`; trượt thì không phát hành. Nhánh `thu/**` chạy cùng quy trình nhưng không đẩy. Xem kết quả: https://github.com/DungHoang1985/ai-da-nguon/actions

Xem bản mới trong Office: đặt con trỏ trong khung add-in rồi Ctrl+F5. Số phiên bản hiện ở đầu khung.

Kiểm trong Office bằng mã (không cần điều khiển màn hình, dùng được cả khi anh không ngồi trước máy): `kiem-thu/cdp.mjs`, cách dùng ở đầu tệp và ở mục Bài học của `tasks/HANDOFF.md`; các bài thử mẫu ở `kiem-thu/cdp-bai/`. Mở và đóng Office thử: `kiem-thu/mo-office.ps1` (tự dừng nếu Office đang mở tài liệu không phải tệp thử). **Thử mã mới trước khi phát hành**: `npm run build` rồi `node kiem-thu/cdp.mjs serve 9229 dist` (chạy nền) nạp bản cục bộ vào khung add-in, không đẩy gì lên `main`. Tệp Word thử: `kiem-thu/docx-thu.mjs` (thăm dò đánh số), `kiem-thu/hop-dong-thu.mjs` (hợp đồng 50 trang có sửa đổi và nhận xét của "Đối tác"), `kiem-thu/soat-thu.mjs` (hợp đồng sạch / cài sẵn 9 lỗi / phiên bản khác lệch 10 chỗ / công văn lề sai, có chú thích, đầu và chân trang, bookmark), `kiem-thu/pdf-quet-thu.mjs` (PDF quét 3 trang).

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
- Khối suy nghĩ của model (Anthropic `thinking`, OpenRouter `reasoning_details`) nằm trong `raw` của tin nhắn trợ lý và được gửi lại nguyên văn, đúng thứ tự; chỗ nào buộc phải đổi tin nhắn cũ (tóm tắt, rút gọn khi lưu) thì bỏ khối suy nghĩ đi kèm. Tham số suy nghĩ không được nằm trong lời nhắc hệ thống, bộ công cụ hay tin nhắn.
- Kỹ năng dựng sẵn: không ghi số hiệu văn bản pháp luật (có kiểm thử chặn).
- Lịch sử hội thoại gửi model là **chỉ thêm**: không sửa, không cắt tin nhắn cũ giữa hai lần tóm tắt; lời nhắc hệ thống không chứa thứ hay đổi (ngày giờ, trạng thái tài liệu). Sửa khác đi là mất bộ nhớ đệm lời nhắc (tốn tiền) — có kiểm thử chặn trong `test/unit/dot0b.test.ts`.
- Tài liệu của nhà cung cấp và của Office.js thay đổi nhanh: đối chiếu tài liệu chính thức hiện hành trước khi viết mã gọi API, không viết theo trí nhớ.
- Không thêm máy chủ riêng, không đổi khung giao diện, không viết lại từ đầu (đã cân nhắc và loại ở mục A của kế hoạch).
- Kho GitHub là **công khai**: không đưa key, token, thông tin riêng tư vào bất kỳ tệp nào.
- Các bài học Office.js đã trả giá nằm ở cuối `tasks/HANDOFF.md`; đọc trước khi đụng tới Word, PowerPoint hoặc WebView2.
- Dữ liệu add-in (key, lịch sử, kỹ năng) nằm trong thư mục `Wef` của Office, tách theo tài khoản Office: không xóa thư mục này trong bất kỳ tệp cài, gỡ hay bước dọn nào.
