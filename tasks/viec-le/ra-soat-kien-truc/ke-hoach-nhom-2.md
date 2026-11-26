# Kế hoạch nhóm 2 — rà soát kiến trúc (lập 26/11/2026, bản 2 sau phản biện; anh duyệt 26/11: làm từng bước, hỏi trước mỗi lần phát hành; bước 4b làm, Chạy thật ghi "chưa kiểm được" tới khi có nguồn)

Nguồn: bảng "Phát hiện để lại" của `ra-soat-kien-truc-2026-11-20.md` (B7, B8, B9, B13) và mục 1–2 "Ý tưởng". Đo lại trên `main@ecedf77` (v4.13.3). Bản 1 đã qua một tác tử phản biện (model khác, chỉ đọc mã); bản này sửa theo phản biện — ghi ở cuối.

## Hiện trạng đo được

| Mục | Hiện trạng 26/11 |
|---|---|
| B7 | `Chat.tsx` 1.815 dòng. `send` (dòng 779) / `prepareAndRun` (797): tải danh mục rồi dựng vòng xoay Miễn phí, xử lý tệp đính kèm (PDF, ảnh, nhắc riêng tư), `groupHints`, khối `onEvent` (~100 dòng: cảnh báo "nói đã sửa mà không ghi"). `receive` (536): nhận lệnh ghi chéo, hẹn giờ 1 giây. Không có kiểm thử máy cho các phần này. |
| B8 | Thông tin một công cụ rải nhiều chỗ: spec, nhãn bước (`tools.ts`), nhóm (`toolgroups.ts`), rủi ro (`RISK` Word trong closure, `RISKY_NAMES` Excel), `UNDO_LABEL` Word, bước Tự kiểm tra. |
| B9 | 66 chỗ so sánh theo chuỗi mã nguồn (`"openrouter"`, `"openai"`, `"mistral"`, `"gemini"`…) trong `src/core`, `src/office`, `src/taskpane` (đếm mọi so sánh với chuỗi mã; bản 1 ghi 43 vì chỉ đếm lõi). |
| B13 | `wordTools.ts` 2.482 dòng: ~480 dòng hàm cấp mô-đun dùng chung + `createWordTools` là một closure ~1.500 dòng giữ trạng thái chung (`lastSnap`, hàng đợi ghi, `tracked`, `UNDO_LABEL`, `RISK`). |
| Lỗi có thật | `receive` chỉ kiểm `busyRef`, không kiểm `sendingRef` (`Chat.tsx:538`): trong khoảng giữa bấm Gửi và lúc khung chuyển "đang chạy" (đọc vùng chọn, tải danh mục tới 8 giây ở Miễn phí), một lệnh ghi chéo có thể chạy chen vào tài liệu. |

## Thứ tự — mỗi bước một bản phát hành riêng

| Bước | Việc | Nghiệm thu (đo được) |
|---|---|---|
| 0 · 4.13.4 | **Vá ghi chéo chen ngang** (`receive` kiểm thêm `sendingRef`) + **kiểm thử golden**: với mỗi ứng dụng, lưu vào kho đúng thứ add-in gửi model và quyết định bộ nhớ đệm — danh sách công cụ `{name, description, parameters}` theo thứ tự (qua đúng hàm dựng trong `llm.ts`), lời nhắc hệ thống Trả phí / Miễn phí, `groupsNote`, `visibleSpecs` cho từng tập nhóm, công cụ chung (`load_tools`…); và thân yêu cầu (địa chỉ, tiêu đề, thân) cho openrouter, openai, groq, mistral, anthropic, gemini × (thường, tìm web, PDF, Suy nghĩ kỹ) | Golden sinh từ 4.13.3 và đạt; ca kiểm thử cho vá `sendingRef`; kiểm thử máy, lint; Tự kiểm tra 3 ứng dụng |
| 1 · 4.14.0 | **B9** — trường `quirks` trong `Provider` (như `dialect`) thay so sánh mã nguồn ở lõi: `llm`, `thinking`, `history` (cả nhóm `mistral`), `rotation`, `modes`, `catalog`, `models`, `types`, `pdfPages`. Giữ nguyên giá trị khóa thời gian nghỉ đã lưu (`cooldownKey`). Nguồn tự khai: không có `quirks` | Golden bước 0 giữ y nguyên từng byte; kiểm thử quét mã cấm `.id === "…"` / `providerId === "…"` ngoài `providers.ts` (danh sách ngoại lệ có chú thích: `oauth.ts`, `credits.ts`, giao diện gắn đúng một nhà cung cấp); kiểm thử khóa nghỉ không đổi; Tự kiểm tra 3 ứng dụng; "Kiểm tra từng nguồn" với các key đang chạy được, so với kết quả 4.13.3 |
| 2a · 4.14.1 | **B13 phần 1** — tách hàm cấp mô-đun ra `wordCore.ts`; `wordTools.ts` vẫn là điểm tái xuất (`WORD_SPECS`, `WORD_SYSTEM_PROMPT`, `reasonOf`, `createWordTools`…) | Golden Word y nguyên (cả thứ tự công cụ); không vòng import (kiểm thử đồ thị import); Tự kiểm tra Word 43/43, các bước Hoàn tác đạt |
| 2b · 4.14.2 | **B13 phần 2** — tách các nhóm công cụ trong closure (đọc, sửa, nhận xét, bảng) theo khuôn `wordReview`/`wordAdmin` (nhận phụ thuộc qua tham số); lớp Hoàn tác + hàng đợi ghi ở lại làm lõi điều phối | Như 2a; `wordTools.ts` dưới 900 dòng |
| 3 · 4.14.3 | **B8** — sổ công cụ một nguồn: nhãn bước, nhóm, rủi ro, nhãn Hoàn tác khai cạnh spec; `tools.ts`, `toolgroups.ts` đọc từ đó. Word → Excel → PowerPoint. Không đổi tập công cụ, thứ tự, nhãn nhóm | Golden 3 ứng dụng y nguyên; ảnh chụp nhãn bước + câu hỏi rủi ro y nguyên; kiểm thử chặn công cụ thiếu khai báo; Tự kiểm tra 3 ứng dụng |
| 4a · 4.15.0 | **B7 phần 1** — đưa ra `src/core`: xử lý tệp đính kèm phần thuần (trả danh sách thông báo theo thứ tự, ảnh đính kèm truyền vào làm tham số), `groupHints`, xử lý một lệnh ghi chéo (hàm thuần; hẹn giờ, trạng thái mỗi lần chạy effect ở lại `Chat.tsx`). Gộp B5 (chụp tài liệu 30 giây, cùng vùng mã) | Kiểm thử máy cho từng hàm (cả ca "Làm tiếp" không có tệp, quá tổng dung lượng tệp, thứ tự thông báo); golden y nguyên; Tự kiểm tra 3 ứng dụng; ghi chéo Word → Excel thử trong Office; Chạy thật Trả phí 1 ca Word có tệp đính kèm |
| 4b · 4.15.1 | **B7 phần 2** — dựng vòng xoay Miễn phí ra lõi, kết quả `{rotation} | {lỗi, trả lại chữ đã gõ}` | Kiểm thử máy với nguồn giả: hủy giữa chừng, lỗi dựng vòng xoay, hết 8 giây chờ, nguồn tải danh mục lỗi; `send` + `prepareAndRun` dưới 250 dòng; Chạy thật Miễn phí khi có nguồn chạy được (hiện chưa có → ghi "chưa kiểm được") |

Mỗi bước: tiêu chí trước khi làm; tác tử phản biện (được đưa golden và danh sách rủi ro) trước khi phát hành; Tự kiểm tra trên bản phát hành; HANDOFF.

## Không làm trong nhóm 2

- Đổi giao diện, hành vi công cụ, lời nhắc hệ thống.
- Khối `onEvent` (cảnh báo "nói đã sửa mà không ghi") và gom 4 đường gọi model: để sau bước 4b, xét lại khi đã có golden.
- B10 phần 2 (nạp công cụ theo ứng dụng): tách hẳn, không gộp vào bước 4.

## Phản biện bản 1 → đã sửa

| Phản biện | Xử lý |
|---|---|
| Ảnh chụp `WORD_SPECS` không đủ giữ bộ nhớ đệm (thứ tự, lời nhắc, nhóm, công cụ chung, thân yêu cầu) | Bước 0: kiểm thử golden làm trước mọi bước |
| B13 không phải "chỉ dời mã": closure 1.500 dòng, hàm dùng chung cần chỗ ở, dễ vòng import | Tách 2a (`wordCore`) / 2b (nhóm); giữ điểm tái xuất; kiểm đồ thị import, Hoàn tác |
| B9 sót `catalog`, `models`, `types`, `pdfPages`, nhóm `mistral`; khóa nghỉ lưu trong máy theo chuỗi mã nguồn | Mở rộng phạm vi; giữ khóa, thêm kiểm thử; kiểm thử quét mã cụ thể |
| Lỗ hổng `sendingRef` trong `receive` (có thật, đã kiểm mã) | Vá ở bước 0 |
| B7 quá lớn, nhiều tác dụng phụ, thứ tự thông báo, "chỉ giữ trạng thái" là quá lời | Tách 4a / 4b; mục tiêu số dòng; kiểm thử thứ tự thông báo; `onEvent` để sau |
| B5 cùng vùng mã | Gộp vào 4a |
| Tiêu chí chưa đo được | Viết lại thành kiểm thử / số đo cụ thể |

## Bước 0 — đã làm (26/11/2026, 4.13.4)

- Vá ghi chéo chen ngang: `receiveBlocked` (`crossapp.ts`) xét thêm cờ đang chuẩn bị gửi; phản biện bước 0 bắt thêm kẽ 80–140 ms lúc chờ giành lệnh → `receive` bật cờ bận ngay khi lấy được lệnh, trước `await claimCommand`.
- Golden `test/unit/golden-gui-model.test.ts` (75 ca, sinh từ mã 4.13.3 chưa sửa): 3 ứng dụng × Trả phí / Miễn phí (thân đầy đủ); 12 nguồn (6 thường + openrouter-claude, deepseek-anthropic, minimax, kimi, 2 nguồn tự khai) × 5 tình huống; 6 nguồn có lịch sử gọi công cụ; nhóm công cụ Miễn phí từng nhóm / mọi nhóm của 3 ứng dụng. `.gitattributes` giữ LF. Còn thiếu (ghi để biết): bước thứ hai trong cùng lượt, khối bộ nhớ / chỉ dẫn riêng, ca gửi lại không tham số suy nghĩ.

## Bước 1 — đã làm (26/11/2026, 4.14.0)

- `ProviderQuirks` (`providers.ts`): 18 trường, gắn đúng các nguồn đã có hành vi đó; `quirksOf(p)` tra lại nguồn dựng sẵn theo mã khi đối tượng không mang `quirks` (vòng xoay, kiểm thử); nguồn tự khai luôn `{}`.
- Thay mọi so sánh mã nguồn ở `history`, `llm`, `thinking`, `rotation`, `catalog`, `models`, `modes`, `pdfPages` và 3 chỗ giao diện (`Chat`, `ModelPicker`, `ProviderCheck`). Còn được phép: `providers.ts`, `credits.ts`, `oauth.ts` (API riêng của OpenRouter). Thẻ `raw.kind` của lịch sử đã lưu giữ nguyên.
- `test/unit/b9-quirks.test.ts`: quét `src/core` + `src/taskpane` (so sánh xuôi, ngược, `case`, biểu thức chính quy trên mã nguồn, `[...].includes(id)`, tập mã nguồn — đã cài thử từng dạng, đều bị bắt); bảng kỳ vọng đủ mọi trường mọi nguồn; khóa thời gian nghỉ, mốc đặt lại hạn mức giữ y nguyên.
- Nghiệm thu: golden 75 ca y nguyên; kiểm thử máy 1.050/1.050; Tự kiểm tra bản cục bộ 43/44/27; "Kiểm tra từng nguồn" 4.13.4 và 4.14.0 cùng 5 nguồn đạt, 5 nguồn lỗi tài khoản (Gemini, Groq, DeepSeek ×2, OpenAI). Phản biện: không lỗi nghiêm trọng; đã làm: so sánh dạng `/^openrouter/` còn sót (`pdfPages`, thành đặc điểm `relay`), 3 chỗ giao diện, mở rộng phép quét, bảng kỳ vọng đủ trường, kiểm mốc Gemini chặt.

## Bước 2a, 2b, 4a, 4b — đã làm (26/11/2026, gộp phát hành; bước 3 gác lại)

- **2a** `wordCore.ts`: 29 hàm / hằng / kiểu cấp mô-đun dời nguyên văn. **2b**: hàm không giữ trạng thái của closure + hằng số vào `wordCore.ts`; 23 nhánh `case` sang `wordRead.ts` (7), `wordEdit.ts` (10), `wordComments.ts` (6) theo khuôn `wordReview`. `wordTools.ts` 2.482 → 853 dòng (đạt mục tiêu < 900). Phản biện: đúng là dời thuần (so từng dòng, từng nhánh). Thêm kiểm thử: không import vòng (đã cài thử một vòng — bị bắt); mỗi công cụ Word có đúng một nơi xử lý.
- **3 (B8) — gác lại, chờ anh quyết**: ba lỗi "quên khai" đã có kiểm thử A8 chặn (nhãn bước, Tự kiểm tra, nhóm chỉ nêu công cụ có thật); nhóm công cụ ở `src/core` (không được phụ thuộc Office) nên không chuyển sang cạnh spec ở `src/office` được mà không phá ranh giới lõi. Phần còn rải: rủi ro (`RISK`, `RISKY_NAMES`, `PPT_RISKY`) và `UNDO_LABEL` — không chặn chung bằng máy được (không phải công cụ nào cũng cần).
- **4a** `core/turnPrep.ts`: `prepareAttachments`, `groupHintsFor`, `handleIncoming`; B5 (`shouldShare`, `clearSnapshot`, bản quá hạn bị gỡ). **4b** `prepareFreeRotation` trong `rotation.ts`. Kiểm thử 19 ca (nguồn giả: Dừng, lỗi, treo, tải lỗi).
- **Không đạt**: `send` + `prepareAndRun` còn 350 dòng (mục tiêu < 250) — phần còn lại là lời gọi `runTurn` và khối `onEvent`, đúng phần kế hoạch ghi "để sau"; mục tiêu đặt lệch với chính phạm vi đó.
- **Nghiệm thu trong Office (bản cục bộ v4.14, desktop ẩn)**: Tự kiểm tra Word 43/43, Excel 44/44, PowerPoint 27/27; Chạy thật Trả phí (Claude Haiku 5.5) hỏi về tệp đính kèm — trả lời đúng ngày chỉ có trong tệp; Chạy thật Miễn phí (OpenRouter Nemotron free) sửa Điều 5 — cổng chặn lần ghi đầu, soát rồi sửa bằng Track Changes, đọc lại (model vẫn không nói căn cứ luật chưa kiểm được — điểm yếu cũ); ghi chéo Excel → Word — hỏi xác nhận, ghi đúng một đoạn có theo dõi, mục "Vừa ghi…" có Hoàn tác, Hoàn tác về đúng như trước. Golden y nguyên qua mọi bước; kiểm thử máy 1.071/1.071.
- Phản biện 4a, 4b: không lỗi nghiêm trọng; **sửa**: B5 bản đầu chỉ chụp khi khung khác *được kết nối* đang mở — tắt "Kết nối với Excel" ở Word làm Excel mất đọc Word (trước đó vẫn đọc được). Nay chụp khi có bất kỳ khung ứng dụng khác mở (`othersOpen`, không lọc theo công tắc). Trả lời khung gửi gọi trong khối thử như cũ (trả lời lỗi → gỡ phần đã ghi). Chạy lại trong Office: Excel đọc Word đúng tiêu đề; ghi chéo + Hoàn tác đạt. Để biết: hai cửa sổ cùng ứng dụng dùng chung khóa chụp — đóng một khung gỡ bản chụp, tối đa 30 giây sau khung kia chụp lại.
- Gộp phát hành thành **4.15.0** (B9 + B13 + B5 + B7). Kiểm thử máy 1.072/1.072.
