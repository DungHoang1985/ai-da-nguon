# Rà soát kiến trúc – 20/11/2026 (trên `main@8546999`, bản 4.11.0)

Anh yêu cầu: "Audit kiến trúc dự án và hoàn thiện. Nếu có ý tưởng mới hãy đề xuất."
Cách làm: đọc cấu trúc, đồ thị import, chạy kiểm thử; ba lượt soát độc lập chỉ đọc (lõi tác tử · lớp công cụ Office · giao diện, đóng gói, bảo mật); mọi mục "cao" được đối chiếu lại với mã trước khi sửa. Nhánh: `claude/project-architecture-audit-38106c`.

## Kết luận

Kiến trúc tổng thể vững: lõi `src/core` không phụ thuộc Office hay giao diện, không có vòng import (86 tệp, 239 cạnh); các quy tắc "không làm trái" của CLAUDE.md (lịch sử chỉ thêm, lời nhắc hệ thống cố định, `dialect`, chế độ Miễn phí không gọi model trả tiền, Track Changes, `expect`, lớp bọc Excel) đều được giữ — trừ một chỗ (mục A1). Điểm yếu chính: (1) thông tin của một công cụ nằm rải ở 6 chỗ, khớp nhau bằng tay; (2) `Chat.tsx` giữ logic điều phối không kiểm thử được; (3) token ghi của CI lộ cho mọi gói npm; (4) một số hộp hỏi Excel / PowerPoint hứa "có thể Hoàn tác" khi không có.

## Tiêu chí nghiệm thu của phiên này · kết quả

| Tiêu chí | Kết quả | Bằng chứng |
|---|---|---|
| Có bảng phát hiện kèm tệp:dòng, đã đối chiếu mã | Đạt | Bảng dưới; mục cao đã đọc lại mã |
| Chỉ sửa chỗ an toàn, kiểm được ngoài Office, không đổi hành vi ngoài phạm vi | Đạt | 9 tệp mã, 1 tệp CI; diff trên nhánh |
| Kiểm thử máy đạt, có kiểm thử mới chặn từng chỗ sửa | Đạt | `npm test` 711/711 (701 + 10 mới ở `test/unit/kien-truc.test.ts`); `npx tsc --noEmit` sạch |
| Nút Chèn Excel không ghi công thức trong Office | Chưa kiểm được | Logic chuyển sang `core/table.ts` và có kiểm thử máy; chưa bấm thử trong Excel |
| Quy trình phát hành chạy với token chỉ ở bước đẩy | Chưa kiểm được | YAML đọc được; cần đẩy lên nhánh `thu/**` để chạy thử (đẩy lên kho là việc cần anh duyệt) |

## Đã sửa trong phiên

| # | Mục | Tệp | Kiểm |
|---|---|---|---|
| A1 | **Cao.** Sau khi tóm tắt và khi lưu, khối suy nghĩ OpenRouter (`reasoning_details`) không bị bỏ — trái quy tắc khối suy nghĩ của CLAUDE.md; model Claude qua OpenRouter có thể trả 400, mất một lần gọi và tắt "Suy nghĩ kỹ" ở mọi lượt sau | `agent/types.ts` (`hasThinkingRaw` dùng chung), `agent/compaction.ts`, `conversations.ts` | 3 kiểm thử (giữ raw Gemini) |
| A2 | **Cao.** Nút Chèn của Excel ghi thẳng chữ AI: dòng "=WEBSERVICE(…)" thành công thức (đường rò dữ liệu nếu AI bị nội dung tài liệu dẫn dắt) | `core/table.ts` (`answerToExcelCells`, `cellAsText` chuyển về lõi), `office/host.ts`, `office/crossExec.ts` (xuất lại) | 3 kiểm thử |
| A3 | **Cao.** CI: `actions/checkout` để token ghi trong `.git/config` suốt `npm ci` và đóng gói | `.github/workflows/phat-hanh.yml`: `persist-credentials: false`, token chỉ ở lệnh `git push` | YAML đọc được; chưa chạy thử |
| A4 | TB. Lượt tóm tắt không truyền `onRetry`, `retryDelaysMs`, `sleep`; lượt soát lập luận không truyền `onRetry` → chờ thử lại tới 23 giây mà khung không báo | `agent/compaction.ts`, `agent/turn.ts` | kiểm thử cũ đạt |
| A5 | Thấp. Lỗi giữa luồng không đóng kết nối HTTP | `agent/llm.ts` (`ctl.abort()` trong `finally`) | kiểm thử cũ đạt |
| A6 | Thấp. Tệp gửi qua key Cohere (có gói dùng thử) thiếu dòng nhắc riêng tư | `pdfPages.ts` | 1 kiểm thử |
| A7 | Thấp. Mã chết `readSSE` | `sse.ts` | tsc |
| A8 | Chặn lệch "sổ công cụ": mọi công cụ của ba ứng dụng phải có nhãn bước, có bước Tự kiểm tra, không trùng tên; nhóm chỉ nêu công cụ có thật | `test/unit/kien-truc.test.ts` | 3 kiểm thử (hiện 87 công cụ khớp hết) |

## Phát hiện để lại (chờ anh chọn)

Sắp theo mức; "Office" = cần Tự kiểm tra hoặc Chạy thật trong Office để nghiệm thu.

| # | Mức | Phát hiện | Chỗ | Cách sửa đề xuất | Office |
|---|---|---|---|---|---|
| B1 | Cao | Hộp hỏi Excel nói "(có thể Hoàn tác)" nhưng vùng > 10.000 ô không được chụp để Hoàn tác; `excel_clear_range` vẫn xóa, `excel_sort_range` không hỏi | `excelCore.ts:6,28`, `excelAdvanced.ts:220-224`, `excelTools.ts:470-490` | Hàm `canUndo(range)` dùng chung cho kiểm rủi ro và `snapshot`; câu hỏi nói đúng; kết quả báo khi không có Hoàn tác | Có |
| B2 | TB | Gửi trùng: `send` chờ `readSelection()` trước khi khóa `busy` → Enter lần hai chạy hai lượt song song trên cùng lịch sử; nút Tự kiểm tra chạy được khi AI đang chạy | `Chat.tsx:762,838`, `Settings.tsx:423` | Khóa `busyRef` ngay đầu `send`; truyền `busy` sang Cài đặt | Có |
| B3 | TB | Excel: không kiểm được ghi đè thì cho chạy (`.catch(() => null)`), Word thì hỏi | `excelAdvanced.ts:221`, `excelModel.ts:418,438,453` | Theo cách của Word: không kiểm được thì hỏi | Có |
| B4 | TB | PowerPoint xóa slide / hình vẫn chạy khi không chụp được bản Hoàn tác | `pptAdvanced.ts:115,651-661,755-765` | Không chụp được thì dừng trước khi xóa | Có |
| B5 | TB | Ảnh chụp tài liệu cho ghi chéo đọc mỗi 30 giây kể cả khi không có ứng dụng khác mở, nằm lại trong localStorage vô thời hạn | `Chat.tsx:463-468`, `share.ts:24-35` | Chỉ chụp khi có khung khác đang mở; xóa khi `pagehide`; `readSnapshot` xóa bản quá hạn | Một phần |
| B6 | TB | Hàng đợi chạy lần lượt lệnh ghi chỉ có ở Word; Excel, PowerPoint có thể chạy xen khi công cụ quá hạn chờ vẫn chạy ngầm | `wordTools.ts:2055-2062` | Hàng đợi cấp mô-đun bọc trong `getHostTools` cho cả ba | Có |
| B7 | TB | `Chat.tsx` (1.670 dòng) tự dựng vòng xoay Miễn phí, định tuyến PDF, `groupHints`, nhận lệnh ghi chéo → quy tắc Đợt 9–11 nửa nằm ở chỗ không có kiểm thử | `Chat.tsx:484-563,778-813,857-916,949-958` | Tách `prepareFreeRotation`, `prepareTurn`, `receiveOnce` về `src/core`, Chat chỉ giữ state | Có (một lượt) |
| B8 | TB | Thông tin một công cụ rải 6 chỗ (spec, nhãn, nhóm, `RISK`/`RISKY_NAMES`, `UNDO_LABEL`, selftest) | `tools.ts`, `toolgroups.ts`, `excelTools.ts:505`, `pptTools.ts:631`, `wordTools.ts:390-407,2027,2160` | Khai `risky`, nhãn, nhóm ngay cạnh spec; A8 đã chặn phần nhãn / selftest / nhóm | Không |
| B9 | TB | Nguồn kiểu OpenAI rẽ nhánh theo `provider.id` ở 8 tệp (tập `{openrouter, openai}` khai lặp 3 nơi) | `llm.ts:308-320,633`, `thinking.ts:100,113`, `history.ts`, `rotation.ts`, `modes.ts` | Thêm trường `quirks` vào `Provider` như `dialect`, làm dần | Không |
| B10 | TB | Gói `taskpane.js` 1,15 MB: Cài đặt nhập tĩnh `selftest*` (~120 KB), nạp công cụ cả ba ứng dụng (~140–190 KB) | `Settings.tsx:7`, `office/tools.ts:3-5` | `import()` khi bấm Tự kiểm tra; nạp công cụ theo ứng dụng | Có |
| B11 | TB | Danh sách hội thoại có thể mất mục khi hai khung cùng lưu; IndexedDB báo xong ở `onsuccess` chứ không ở `oncomplete`; mở DB lỗi một lần là hỏng mãi | `conversations.ts:97-109`, `db.ts:30-42` | Một giao dịch đọc–sửa–ghi; resolve ở `oncomplete`; lỗi thì chuyển `MemoryKV` và báo | Một phần |
| B12 | TB | Nguồn gốc `dunghoang1985.github.io` dùng chung cho mọi trang Pages của tài khoản: trang khác cùng tài khoản đọc được key | `webpack.config.js:10` | Không bật Pages cho kho khác của tài khoản; lâu dài: tên miền riêng | Không |
| B13 | Thấp | `wordTools.ts` 2.299 dòng | — | Tách `wordUndo`, `wordRead`, `wordEdit`, `wordComments`, `wordTables` theo khuôn `createXTools` | Có (Word 39/39) |
| B14 | Thấp | Tự thử lại sau khi model đã sinh suy nghĩ / tham số công cụ tính tiền thêm | `llm.ts:621-640` | Cân nhắc: tự thử lại vẫn tiện hơn bắt người dùng gửi lại; **đề xuất giữ nguyên** | — |
| B15 | Thấp | Lint không chạy được (3.604 lỗi, chủ yếu CRLF; thiếu `react-hooks`), CI không chạy lint | — | `eslint.config.mjs` riêng, `endOfLine: auto`, chạy ở CI chế độ báo | Không |
| B16 | Thấp | webpack dịch `.ts` hai lần (ts-loader rồi babel-loader); 5 devDeps không dùng | `webpack.config.js:43-54` | Một rule ts-loader; gỡ devDeps | Có (build + Tự kiểm tra) |
| B17 | Thấp | Không có CSP ở `taskpane.html`, `dang-nhap.html` | — | `script-src 'self' <CDN Office>`, `object-src 'none'` | Có |
| B18 | Thấp | Mỗi token trả về vẽ lại toàn bộ hội thoại | `Chat.tsx:1299-1320` | `React.memo` cho từng mục | Có |
| B19 | Thấp | Word dưới WordApi 1.6: lệnh ghi không có Hoàn tác | `wordTools.ts:2066` | Từ chối ghi như khi dưới 1.4 | Không |
| B20 | Thấp | Kho công khai còn đường dẫn máy (`C:\Users\…`, `D:\…`) và một kế hoạch nội bộ | `tasks/ban-giao-phan-phoi-cong-ty-va-che-do-web.md`, 3 tệp `soat-ma-*` | Đường dẫn tương đối; chuyển kế hoạch nội bộ ra ngoài kho | Không |
| B21 | Thấp | `--remote-allow-origins=*` khi mở Office thử | `kiem-thu/mo-office.ps1:95` | Chỉ cho `http://127.0.0.1:<cổng>` | Có |
| B22 | Thấp | `customProviders` chỉ nạp như tác dụng phụ của `Storage.getKeys()` | `providers.ts:322`, `storage.ts:126` | Nạp tường minh lúc khởi động | Không |
| B23 | Cần anh | PowerPoint ghi đè chữ ô bảng / hình có sẵn không hỏi (Word, Excel thì hỏi) | `pptDesign.ts:68-84` | Anh quyết quy tắc "ghi đè hỏi trước" có áp cho PowerPoint không | — |

Đã kiểm, không thấy lỗi: không có XSS (không `innerHTML`/`dangerouslySetInnerHTML`/`eval`); OAuth PKCE đúng quy tắc Đợt 2; không có key thật trong kho; 84/87 công cụ có bước Tự kiểm tra (3 còn lại là công cụ dùng chung không gọi Office.js); Office.js không dùng proxy sau khi ngữ cảnh đóng.

## Ý tưởng mới đề xuất

1. **Sổ công cụ một nguồn (`defineTool`)**: mỗi công cụ khai một lần: spec, nhãn bước, nhóm, `risky`, nhãn Hoàn tác, bước Tự kiểm tra. Gộp B8, chặn cả lớp lỗi "quên thêm vào bảng". Làm dần theo từng ứng dụng.
2. **Lõi điều phối lượt hoàn chỉnh** (B7 + gom 4 đường gọi model về một hàm `call` dựng trong `runTurn`): giao diện chỉ còn state; mọi quy tắc Đợt 9–11 kiểm thử máy được.
3. **Ước tính tiền mỗi lượt**: khung đã có số token mỗi tin; ghép bảng giá (OpenRouter trả `pricing` trong danh mục, đang đọc ở `models.ts:73`) để hiện "≈ 0,02 USD" cạnh mỗi câu trả lời và tổng theo hội thoại — giúp chọn model ở chế độ Trả phí.
4. **Xem trước thay đổi Excel**: trước lệnh ghi lớn, hiện bảng "ô · cũ → mới" (đã có `diffCells` trong `auditlog.ts`) để bấm Đồng ý, tương tự Track Changes của Word.
5. **Phát hành an toàn hơn**: tách CI thành job đóng gói (chỉ đọc) và job đẩy (không cài gói); thêm bước Tự kiểm tra bằng `kiem-thu/cdp.mjs` thành danh sách kiểm trước phát hành có máy ghi số.

## Nhóm 3 — làm ở 4.13.3 (26/11/2026)

| Mục | Kết quả | Đo / bằng chứng |
|---|---|---|
| B10 | Mã Tự kiểm tra (`selftest*`) nạp bằng `import()` khi bấm nút; lỗi nạp hiện thành một bước lỗi. **Chưa làm**: tách công cụ theo ứng dụng (cần nạp công cụ bất đồng bộ trong `Chat.tsx`, đụng ghi chéo) | Gói `taskpane` 1.293.744 → 1.145.405 byte; tệp riêng `tu-kiem-tra` 152.578 byte. Đo theo bản đồ nguồn: bộ biểu tượng Fluent đã được lọc ở bản cuối, không phải nguyên nhân |
| B16 | Chỉ ts-loader; gỡ `babel.config.json`, `@babel/*`, `babel-loader`, `source-map-loader`, `@types/webpack` | Gói không đổi kích thước (Babel 8 mặc định đích hiện đại) — chỉ bớt một lượt dịch |
| B18 | `Rich`, `Linked` bọc `React.memo`; `onNav` dùng `useCallback([host])`, đặt trước hai chỗ `return` sớm | Lần đầu đặt hook sau `return` sớm (sẽ vỡ khung khi chọn chế độ) — bắt kịp, nay lint chặn |
| B15 | `eslint.config.mjs` riêng: lỗi JS cơ bản + `react-hooks/rules-of-hooks` (thư mục giao diện) mức lỗi; quy tắc Office.js mức cảnh báo (chỉ `src/office`); bỏ quy tắc định dạng. CI chạy `npm run lint`, lỗi thì không phát hành | 0 lỗi, 285 cảnh báo. Thử cài một hook sau `return` sớm: lint báo lỗi đúng chỗ |

Tự kiểm tra bản cục bộ 4.13.3: Word 43/43, Excel 44/44, PowerPoint 27/27. Tác tử phản biện (model khác, không thấy lập luận): không lỗi nghiêm trọng; đã làm: khai báo trực tiếp `eslint`, `@typescript-eslint/parser`, `@typescript-eslint/eslint-plugin` (trước chỉ có gián tiếp qua `office-addin-lint`), đồng bộ số bản trong lock, hai chỗ vụn ở `Chat.tsx`.

## Nhóm 2 — làm ở 4.15.0 (26/11/2026)

B9, B13, B7 (phần thuần; khối `onEvent` để sau), B5 xong; **B8 không làm** (anh quyết): lỗi "quên khai" đã có kiểm thử A8 chặn, nhóm công cụ nằm ở lõi không phụ thuộc Office. Chi tiết, nghiệm thu, phản biện: `ke-hoach-nhom-2.md`.
