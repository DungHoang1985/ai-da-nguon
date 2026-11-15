# AI Đa Nguồn — add-in AI cho Word, Excel, PowerPoint

Add-in Office dùng API key của nhiều nhà cung cấp, AI tự thao tác trên tài liệu như add-in Claude chính thức. Dùng cá nhân (anh Dũng, pháp lý – bất động sản). Đang chạy **v3.5**; kế hoạch lên v4 (bản 2) đã duyệt ngày 15/11/2026.

## Đọc trước khi làm bất cứ việc gì

1. `tasks/HANDOFF.md` — trạng thái, quyết định, vấn đề mở, bài học.
2. `tasks/ke-hoach-v4.md` — hiện trạng đã kiểm, phương án đăng nhập, kế hoạch theo đợt (0A, 0B, 1A, 1B, 2, 3, 4) và tiêu chí nghiệm thu. Bằng chứng rà soát: `tasks/ra-soat-2026-11-15/DOC-TRUOC.md`.
3. `tasks/dot-0.md` — việc cụ thể của đợt đang làm.
4. `docs/so-sanh-claude-chinh-thuc.md` — bảng so sánh 55 khả năng với bản chính thức.

## Nguồn mã

Từ 15/11/2026 (bước 1 của `tasks/dot-0.md`), thư mục này là bản làm việc của nhánh `main` — một nguồn duy nhất, gồm cả `docs/`, `tasks/`, `kiem-thu/` và tệp này.

- Kho: https://github.com/DungHoang1985/ai-da-nguon (công khai) — nhánh `main` (mã), nhánh `gh-pages` (bản đang chạy).
- Bản chạy: https://dunghoang1985.github.io/ai-da-nguon/
- Trước khi sửa: `git status` sạch và `git pull` để lấy bản mới nhất.

## Cấu trúc mã (nhánh `main`)

```
src/core/            lõi không phụ thuộc Office, kiểm thử được bằng npm test
  agent/loop.ts      vòng lặp tác tử (gọi model → chạy công cụ → lặp)
  agent/llm.ts       gọi model, gom 3 kiểu API: OpenAI, Anthropic, Gemini
  agent/compaction.ts  tự tóm tắt hội thoại dài
  providers.ts models.ts storage.ts   nhà cung cấp, danh mục model, key lưu cục bộ
  skills.ts skillpack.ts              kỹ năng (SKILL.md), 17 kỹ năng dựng sẵn
  conversations.ts db.ts share.ts attachments.ts refs.ts
src/office/          công cụ Office.js: excel*, word*, ppt*, common, selftest (nút Tự kiểm tra)
src/taskpane/components/   App.tsx, Chat.tsx (điều phối + giao diện), Settings.tsx
test/unit/           mocha, 71 kiểm thử
cai-dat/             tệp .cmd cài và gỡ add-in vào Office (ghi registry WEF\Developer)
```

## Lệnh

```
npm ci                      # .npmrc phải trỏ registry.npmjs.org
npm test                    # 71/71 phải đạt
npx tsc --noEmit -p .       # không lỗi kiểu
npm run build               # ra dist/, manifest tự đổi sang địa chỉ GitHub Pages
```

Phát hành hiện làm tay (chép `dist/` sang nhánh `gh-pages`). Đợt 0 chuyển sang tự động bằng GitHub Actions; sau đó chỉ cần đẩy lên `main`.

Xem bản mới trong Office: đặt con trỏ trong khung add-in rồi Ctrl+F5. Số phiên bản hiện ở đầu khung.

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
- Kỹ năng dựng sẵn: không ghi số hiệu văn bản pháp luật (có kiểm thử chặn).
- Tài liệu của nhà cung cấp và của Office.js thay đổi nhanh: đối chiếu tài liệu chính thức hiện hành trước khi viết mã gọi API, không viết theo trí nhớ.
- Không thêm máy chủ riêng, không đổi khung giao diện, không viết lại từ đầu (đã cân nhắc và loại ở mục A của kế hoạch).
- Kho GitHub là **công khai**: không đưa key, token, thông tin riêng tư vào bất kỳ tệp nào.
- Các bài học Office.js đã trả giá nằm ở cuối `tasks/HANDOFF.md`; đọc trước khi đụng tới Word, PowerPoint hoặc WebView2.
- Dữ liệu add-in (key, lịch sử, kỹ năng) nằm trong thư mục `Wef` của Office, tách theo tài khoản Office: không xóa thư mục này trong bất kỳ tệp cài, gỡ hay bước dọn nào.
