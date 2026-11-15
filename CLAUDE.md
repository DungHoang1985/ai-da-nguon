# AI Đa Nguồn — add-in AI cho Word, Excel, PowerPoint

Add-in Office dùng API key của nhiều nhà cung cấp, AI tự thao tác trên tài liệu như add-in Claude chính thức. Dùng cá nhân (anh Dũng, pháp lý – bất động sản). Đang chạy **v3.7.0**; kế hoạch lên v4 (bản 2) đã duyệt ngày 15/11/2026.

## Đọc trước khi làm bất cứ việc gì

1. `tasks/HANDOFF.md` — trạng thái, quyết định, vấn đề mở, bài học.
2. `tasks/ke-hoach-v4.md` — hiện trạng đã kiểm, phương án đăng nhập, kế hoạch theo đợt (0A, 0B, 1A, 1B, 2, 3, 4) và tiêu chí nghiệm thu. Bằng chứng rà soát: `tasks/ra-soat-2026-11-15/DOC-TRUOC.md`.
3. `tasks/dot-0b.md` — việc và kết quả của Đợt 0B (Đợt 0A: `tasks/dot-0.md`, đã xong). Đợt kế tiếp là 1A: viết `tasks/dot-1a.md` trước khi làm.
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
  agent/llm.ts       gọi model, gom 3 kiểu API: OpenAI, Anthropic, Gemini; tự thử lại, bộ nhớ đệm Anthropic, đếm token
  agent/prompt.ts    lời nhắc hệ thống cố định + quy tắc an toàn; trạng thái tài liệu đi trong tin nhắn của lượt
  agent/compaction.ts  tự tóm tắt hội thoại dài, ngưỡng theo sức chứa model
  providers.ts models.ts storage.ts   nhà cung cấp, danh mục model, key lưu cục bộ
  skills.ts skillpack.ts              kỹ năng (SKILL.md), 17 kỹ năng dựng sẵn
  conversations.ts db.ts share.ts attachments.ts refs.ts
src/office/          công cụ Office.js: excel*, word*, ppt*, common, selftest (nút Tự kiểm tra)
src/taskpane/components/   App.tsx, Chat.tsx (điều phối + giao diện), Settings.tsx
test/unit/           mocha, 150 kiểm thử
cai-dat/             tệp .cmd cài và gỡ add-in vào Office (ghi registry WEF\Developer)
```

## Lệnh

```
npm ci                      # .npmrc phải trỏ registry.npmjs.org
npm test                    # 150/150 phải đạt
npx tsc --noEmit -p .       # không lỗi kiểu
npm run build               # ra dist/, manifest tự đổi sang địa chỉ GitHub Pages
```

Phát hành tự động: đẩy lên `main` là GitHub Actions (`.github/workflows/phat-hanh.yml`) kiểm thử, kiểm kiểu, đóng gói rồi đưa `dist/` lên `gh-pages`; trượt thì không phát hành. Nhánh `thu/**` chạy cùng quy trình nhưng không đẩy. Xem kết quả: https://github.com/DungHoang1985/ai-da-nguon/actions

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
- Lịch sử hội thoại gửi model là **chỉ thêm**: không sửa, không cắt tin nhắn cũ giữa hai lần tóm tắt; lời nhắc hệ thống không chứa thứ hay đổi (ngày giờ, trạng thái tài liệu). Sửa khác đi là mất bộ nhớ đệm lời nhắc (tốn tiền) — có kiểm thử chặn trong `test/unit/dot0b.test.ts`.
- Tài liệu của nhà cung cấp và của Office.js thay đổi nhanh: đối chiếu tài liệu chính thức hiện hành trước khi viết mã gọi API, không viết theo trí nhớ.
- Không thêm máy chủ riêng, không đổi khung giao diện, không viết lại từ đầu (đã cân nhắc và loại ở mục A của kế hoạch).
- Kho GitHub là **công khai**: không đưa key, token, thông tin riêng tư vào bất kỳ tệp nào.
- Các bài học Office.js đã trả giá nằm ở cuối `tasks/HANDOFF.md`; đọc trước khi đụng tới Word, PowerPoint hoặc WebView2.
- Dữ liệu add-in (key, lịch sử, kỹ năng) nằm trong thư mục `Wef` của Office, tách theo tài khoản Office: không xóa thư mục này trong bất kỳ tệp cài, gỡ hay bước dọn nào.
