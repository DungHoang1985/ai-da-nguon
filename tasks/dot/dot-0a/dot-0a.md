# Đợt 0A — Nền phát hành và bốn việc anh nêu (v3.5 → v3.6)

Đã duyệt 15/11/2026. Tiêu chí nghiệm thu: bảng Đợt 0A trong `tasks/ke-hoach-v4.md` mục 5. Lý do từng việc: mục 2 và 3 của cùng tệp. Bằng chứng rà soát: `tasks/luu/ra-soat-2026-11-15/`.

Làm theo thứ tự. Sửa lỗi thì viết kiểm thử trước. Mỗi bước xong chạy `npm test` + `npx tsc --noEmit -p .` rồi mới commit, mỗi bước một commit. Số dòng mã dưới đây là của commit `a38ee20`; kiểm lại trước khi sửa.

## Điều kiện

| Cần | Trạng thái |
|---|---|
| Git, Node.js, quyền đẩy lên kho | Đã kiểm 14/11: Git 2.56, Node 24.14, `git push --dry-run` đạt |
| Điều khiển Word, Excel, PowerPoint trên máy | Chưa thử. Xin anh cấp quyền từng ứng dụng khi tới bước cần. Không có thì mục phải kiểm trong Office ghi "Chưa kiểm được", không ghi Đạt thay |

## Bước 1 — Một nguồn mã duy nhất

Biến `D:\ADD IN` thành bản làm việc của nhánh `main`, giữ nguyên `docs/`, `tasks/`, `kiem-thu/`, `CLAUDE.md`.

1. Chuyển mọi thứ khác ở gốc thư mục (`src/`, `test/`, `assets/`, `manifest.xml`, `package*.json`, `README.md`, `cai-dat.cmd`, tệp cấu hình) vào `_ban-0-cu/`. Chuyển, chưa xóa.
2. `git init -b main` → `git remote add origin https://github.com/DungHoang1985/ai-da-nguon.git` → `git fetch origin` → `git checkout -B main origin/main`.
3. Thêm `_ban-0-cu/` vào `.gitignore`.
4. Rà `docs/`, `tasks/` (cả `tasks/luu/ra-soat-2026-11-15/`), `kiem-thu/`, `CLAUDE.md`: không có key, email, thông tin riêng tư (kho công khai). Xóa `tasks/plan.md`, `tasks/todo.md` (lỗi thời). Rồi commit.
5. `npm ci` → `npm test` (71/71) → `npm run build`.
6. Đạt hết thì xóa `_ban-0-cu/` (anh đã duyệt xóa; hệ thống hỏi quyền thì xin một lần).

Kiểm: `git status` sạch; `manifest.xml` có Id `c95713b4-2c26-42c9-944c-cdbd5f4ff7e9`.

## Bước 2 — Phát hành tự động

`.github/workflows/`: khi đẩy lên `main` chạy `npm ci` → `npm test` → `npx tsc --noEmit` → `npm run build` → đưa `dist/` (không kèm tệp `.map`) lên nhánh `gh-pages`, giữ `.nojekyll`. Kiểm thử trượt thì không phát hành.

Kiểm: một lần chạy xanh và trang Pages đổi theo; một lần cố ý làm trượt kiểm thử trên nhánh thử thì không phát hành.

## Bước 3 — Gỡ phần thừa

| Việc | Ở đâu |
|---|---|
| Xóa `src/core/chat.ts` và kiểm thử của nó | `src/core/`, `test/unit/` |
| Bỏ gói đệm trình duyệt cũ: mục `polyfill` trong `entry` và `chunks`, `core-js`, `regenerator-runtime`, mục tiêu "ie 11", thẻ `X-UA-Compatible` | `webpack.config.js:20,69,97`, `package.json`, `babel.config.json`, `taskpane.html:6` |
| Gỡ thư viện không dùng (tìm kiếm trước khi gỡ từng cái): `es6-promise`, `@fluentui/react-icons`, `less`, `less-loader`, `file-loader`, `os-browserify`, `process`, `acorn`, `office-addin-mock`, `office-addin-test-helpers`, `office-addin-test-server` | `package.json` |
| Sửa tên kho, một nguồn số phiên bản lấy từ `package.json` | `package.json`, `webpack.config.js`, `src/core/version.ts` |
| `npm audit fix` cho `image-size`; kiểm lại biểu đồ PowerPoint | — |

Lưu ý: bố cục mới (bước 6) cần biểu tượng. Nếu dùng `@fluentui/react-icons` thì giữ lại gói này.

## Bước 4 — Lỗi mất kết nối khi đổi nhà cung cấp

Nguyên nhân: `tasks/ke-hoach-v4.md` mục 3.3. Anh không nhớ cặp nguồn nào gây lỗi.

1. **Tái hiện trước khi sửa** (cần điều khiển Office): với mỗi cặp nguồn anh có key, chat một lượt có gọi công cụ ở nguồn A, đổi sang B, gửi tiếp; ghi lại dòng lỗi. Lập bảng cặp nguồn.
2. **Kiểm thử trước:** đưa bảng thử trong `tasks/luu/ra-soat-2026-11-15/ma-tran-doi-nha-cung-cap.txt` thành kiểm thử máy cho lớp chuyển đổi; viết kiểm thử cho phần trạng thái danh mục model (tách khỏi `Chat.tsx` thành hàm thuần).
3. **Sửa trạng thái danh mục** (`Chat.tsx:276–282, 333–363, 416–422`): đổi nguồn thì xóa ngay danh sách và model cũ; có trạng thái đang tải / lỗi kèm lý do; bỏ kết quả tải về muộn của nguồn khác; không cho gửi khi model không thuộc nguồn đang chọn; đang chạy thì việc đổi có hiệu lực từ lượt sau.
4. **Nhớ model theo từng nguồn** (`storage.ts:128–134`): lưu `{nguồn → model}`; nguồn chưa từng dùng thì chọn model gọi được công cụ, không lấy mục đầu theo chữ cái (`models.ts:16–21`).
5. **Lỗi hiện đúng** (`errors.ts:21–28`): chỉ lỗi `fetch` thật mới báo "Không kết nối được"; lỗi khác hiện nội dung thật và có nút chép chi tiết.
6. **Nguồn không gọi được từ add-in** (`providers.ts`): gắn cờ cho SambaNova, NVIDIA (đo 14/11: không trả tiêu đề cho phép); hiện lý do ngay khi chọn. Đo lại bằng nút ở mục 8 trước khi chốt.
7. **Chuẩn hóa lịch sử khi đổi nguồn** (`agent/llm.ts` `buildStepRequest`): đổi mã lời gọi công cụ về dạng nguồn đích chấp nhận (Mistral đúng 9 ký tự chữ-số; Anthropic chỉ chữ, số, `_`, `-`; OpenAI tối đa 40 ký tự), giữ khớp cặp gọi–kết quả; bỏ khối chữ chỉ có khoảng trắng; bỏ ảnh cũ khi model đích không nhận ảnh và ghi chú thay thế.
8. **Nút Kiểm tra nhà cung cấp** (trang Cài đặt): với mỗi key, chạy thật chuỗi "model gọi một công cụ giả → nhận kết quả → trả lời có chứa kết quả". Báo Đạt hoặc lý do bằng tiếng Việt.

## Bước 5 — Enter để gửi

`Chat.tsx:974–976`. Enter gửi; Shift+Enter xuống dòng; Ctrl+Enter vẫn gửi; không gửi khi `isComposing` (bộ gõ đang ghép chữ). Danh sách "/" đang mở: Enter hoặc Tab chọn mục, mũi tên di chuyển, Esc đóng. Đang chạy: Enter không gửi, Esc = Dừng. Sửa chữ gợi ý trong ô nhập. Kiểm bằng Unikey và bộ gõ Telex của Windows.

## Bước 6 — Bố cục "một dòng đầu" (anh đã chọn)

Trước khi dựng: chụp khung add-in Claude chính thức đang cài trên máy để đối chiếu vị trí nút chọn model, lịch sử, hội thoại mới.

1. **Đầu khung một dòng** (`App.tsx`): tên + số phiên bản nhỏ · Hội thoại mới · Lịch sử · Cài đặt. Bỏ `TabList`. Cài đặt thành trang riêng có nút quay lại; thẻ Chat vẫn giữ nguyên trạng thái khi mở Cài đặt (hiện đang giữ bằng `display: none`, `App.tsx:48`).
2. **Ô nhập**: tự giãn 1–6 dòng. Hàng nút bên dưới: "+" (đính kèm tệp, kỹ năng, tìm web, gửi kèm vùng chọn) · mức quyền (Hỏi trước khi sửa / Tự sửa, hỏi khi rủi ro — nối với `askBeforeWrite`, thao tác rủi ro vẫn luôn hỏi) · nhãn model · Gửi / Dừng.
3. **Bảng chọn model** (bung từ nhãn model): nhà cung cấp, ô tìm, lọc Miễn phí, Chỉ model thao tác được, tải lại, ghi chú danh mục. Dùng chung trạng thái đã sửa ở bước 4.
4. **Trang Cài đặt** (`Settings.tsx`): tài khoản và key · model mặc định · quyền và an toàn · chỉ dẫn riêng · kỹ năng · đọc chéo ứng dụng · sao lưu · tự kiểm tra · kiểm tra nhà cung cấp. Chuyển công tắc "AI tự thao tác" vào đây; khi tắt, cạnh ô nhập hiện nhãn "Chỉ trò chuyện".
5. **Bỏ công tắc "Word: ghi sửa đổi bằng Track Changes"** (anh đã chọn): gỡ khỏi `Settings.tsx:125` và `storage.ts`; `wordTools.ts` luôn bật theo dõi trước khi ghi và trả về chế độ cũ sau khi ghi; bản Word không có API theo dõi (dưới WordApi 1.4) thì từ chối ghi, báo rõ.
6. **Màn chào** khi chưa có key nào.
7. Bài học cũ: phần tử trong cột flex phải đặt `flex-shrink: 0` nếu không muốn bị co.

Thêm bước Tự kiểm tra: tắt Track Changes trong Word rồi gọi một công cụ ghi → sửa đổi vẫn được ghi có theo dõi.

## Bước 7 — Sao lưu và khôi phục

Trang Cài đặt: **Sao lưu** gom key, cài đặt, chỉ dẫn riêng, model đã nhớ (localStorage, tiền tố `aidn.`) và kỹ năng, lệnh, lịch sử (IndexedDB qua `db.ts`) thành một tệp JSON có số phiên bản định dạng; **Khôi phục** đọc tệp, hỏi gộp hay thay. Ghi rõ trên giao diện: tệp chứa key, là tệp bí mật.

Thử trước: tải tệp từ khung add-in ra đĩa. Không được thì dùng phương án chép vào bộ nhớ tạm và dán từ bộ nhớ tạm.

Khi mở khung mà kho trống trong khi máy từng có dữ liệu (dấu hiệu đổi tài khoản Office hoặc Office đổi thư mục lưu): hiện lời nhắc khôi phục. Không đụng thư mục `Wef` trong mọi tệp cài, gỡ.

## Bước 8 — Ba lỗi Word nhỏ

Chi tiết và cách sửa: `tasks/luu/ra-soat-2026-11-15/ma-word.md`, phát hiện 3, 11, 15.

1. `word_read_paragraphs` (`wordTools.ts:344–360`, `loop.ts:6,96`): tính ngân sách theo độ dài JSON thật, đặt `next_start` và `truncated` lên đầu kết quả. Rà cùng lỗi ở `word_read_tables`, `word_list_comments`, `word_list_revisions`.
2. `word_list_comments`, `word_list_revisions`: tìm số đoạn bằng định danh đoạn (`uniqueLocalId`, WordApi 1.6) thay cho so chuỗi.
3. Trạng thái văn bản trong lời nhắc bị cắt 6.000 ký tự: đưa vùng đang chọn, chế độ theo dõi, nhận xét lên đầu; cắt danh sách tiêu đề theo phần tử, không cắt chuỗi JSON.

## Bước 9 — Kiểm trong Office và đóng đợt

1. Chạy đủ bảng tiêu chí Đợt 0A; mỗi mục một ảnh hoặc kết quả lệnh. Tự kiểm tra Excel 22/22, Word 14/14 cộng bước mới, PowerPoint 14/14.
2. Nâng phiên bản 3.6.
3. Sửa `docs/so-sanh-claude-chinh-thuc.md` theo mục 7 của kế hoạch.
4. Cập nhật `tasks/HANDOFF.md`; báo bảng Tiêu chí · Kết quả · Bằng chứng.
5. Tạo `tasks/dot/dot-0b/dot-0b.md` cho Đợt 0B trước khi bắt đầu đợt đó; khi đó rà nốt mã lõi tác tử.
