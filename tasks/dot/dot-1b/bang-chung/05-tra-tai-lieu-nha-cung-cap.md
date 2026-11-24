# Tra tài liệu chính thức trước khi viết mã gọi API (15/11/2026)

Nguồn đọc ngày 15/11/2026. Mục ghi "chưa xác minh" là tài liệu không nói rõ; mã phải chịu được cả hai khả năng.

## Anthropic (Messages API) — theo tài liệu đi kèm bộ công cụ Claude API (bản lưu 03/11/2026)

- **PDF**: khối `{"type":"document","source":{"type":"base64","media_type":"application/pdf","data":"<base64>"}}` đặt trong nội dung tin nhắn người dùng, **trước** khối chữ. Giới hạn: 32 MB cho cả yêu cầu, 600 trang (100 trang với model sức chứa 200K). Không cần tiêu đề beta.
- **Suy nghĩ**:
  - Model đời mới (Fable/Mythos 5.x, Opus 5.5, Opus 5, Sonnet 5.5, Sonnet 5, Haiku 5.5): suy nghĩ **bật sẵn** khi không gửi `thinking` (kiểu "adaptive"); `{"type":"enabled","budget_tokens":N}` trả 400. Điều khiển mức suy nghĩ bằng `output_config.effort` (`low|medium|high|xhigh|max`); mặc định `medium` ở Opus 5.5 và Haiku 5.5, `high` ở các model khác.
  - Opus 4.8 / 4.7: không gửi `thinking` thì **không** suy nghĩ; bật bằng `{"type":"adaptive"}`; `budget_tokens` trả 400.
  - Opus 4.6 / Sonnet 4.6: `{"type":"adaptive"}`; mức `xhigh` không có (chỉ `low|medium|high|max`).
  - Haiku 4.5 và model cũ hơn: `{"type":"enabled","budget_tokens":N}`, N ≥ 1024 và nhỏ hơn `max_tokens`; không có `effort`.
  - Chữ suy nghĩ mặc định bị bỏ trống (`display: "omitted"`), khối `thinking` chỉ còn `signature`. Vẫn tính tiền như nhau.
- **Giữ khối suy nghĩ**: gửi lại nguyên vẹn các khối `thinking` / `redacted_thinking` (kèm `signature`) ở mọi yêu cầu sau. Trên Fable 5.1/5.2, Opus 5.5, Sonnet 5.5, Haiku 5.5 khối suy nghĩ gắn với **phần hội thoại đứng trước nó** (lời nhắc hệ thống, bộ công cụ, mọi tin nhắn trước): sửa, bỏ, đổi thứ tự tin nhắn cũ, đổi lời nhắc hệ thống hay bộ công cụ giữa hội thoại thì yêu cầu gửi lại khối đó bị 400 (`Invalid signature in thinking block… bound to a different conversation`) — áp dụng mặc định cho tài khoản tạo từ 31/8/2026; Fable 5.2 áp dụng cho mọi tài khoản. Lịch sử chỉ thêm thì không vướng.
  - Hai cách gỡ khi vướng: (1) bỏ hết khối `thinking` và `redacted_thinking` khỏi lịch sử rồi gửi lại một lần; (2) gửi `thinking: {"type":"adaptive","block_binding":{"prefix_mismatch_behavior":"drop_block"}}` kèm tiêu đề `anthropic-beta: thinking-binding-controls-2026-08-01` để máy chủ tự bỏ khối không còn hợp lệ. `block_binding` chỉ đi cùng `adaptive`; gửi mà thiếu tiêu đề beta thì 400.
  - Đổi `max_tokens`, `output_config.effort`, mốc `cache_control` không làm hỏng khối suy nghĩ.
- Sự kiện luồng: `content_block_start` (khối `thinking` / `redacted_thinking`), `content_block_delta` với `thinking_delta` (chữ) và `signature_delta` (chữ ký).

## Gemini (`generateContent`, v1beta) — aistudio.google.com/docs (trang cũ ai.google.dev chuyển hướng về đây)

- **PDF**: phần `{"inlineData":{"mimeType":"application/pdf","data":"<base64>"}}` (tài liệu REST viết `inline_data` / `mime_type`; bản tham chiếu API và ví dụ JS viết `inlineData` / `mimeType` — mã hiện có của add-in đang dùng dạng sau cho ảnh và chạy được). Tối đa 50 MB hoặc 1000 trang; mỗi trang được đọc bằng thị giác (khoảng 258 token mỗi trang) nên PDF quét đọc được. Trang `…/docs/document-processing?generate-content=true`.
- **Suy nghĩ**: `generationConfig.thinkingConfig`.
  - Dòng 2.5: `thinkingBudget` (−1 = tự động; 2.5 Pro 128–32768; 2.5 Flash 0–24576); **không** nhận `thinkingLevel`.
  - Dòng 3.x và 4: `thinkingLevel` (`minimal|low|medium|high`, riêng gemini-4-hydrogen có `xhigh`); suy nghĩ bật sẵn; ở Gemini 4 gửi cả `thinkingBudget` lẫn `thinkingLevel` thì 400.
  - Gửi `thinkingConfig` cho model không hỗ trợ: chưa xác minh (có thể 400) → gặp 400 thì gửi lại không có `thinkingConfig`.
- **Chữ ký suy nghĩ** (`thoughtSignature`): gửi lại nguyên các `parts` của câu trả lời, không gộp phần — add-in đã làm từ bản 3.x (`raw.parts`).

## OpenRouter (`/api/v1/chat/completions`)

- **PDF** (`openrouter.ai/docs/guides/overview/multimodal/pdfs.md`): phần nội dung `{"type":"file","file":{"filename":"x.pdf","file_data":"data:application/pdf;base64,<b64>"}}`. Tùy chọn `plugins: [{"id":"file-parser","pdf":{"engine":"…"}}]`: `native` (model tự đọc tệp, tính token đầu vào), `mistral-ocr` (2 USD / 1000 trang, khuyên dùng cho bản quét), `cloudflare-ai` (miễn phí, lấy lớp chữ). Không chọn engine: dùng cách đọc gốc của model nếu có, không thì `mistral-ocr` (**có phí**). Kết quả phân tích nằm ở `annotations` của câu trả lời; gửi lại `annotations` ở lượt sau thì không bị phân tích (và tính phí) lần nữa.
- **Suy nghĩ** (`…/guides/best-practices/reasoning-tokens.md`): `reasoning: {"effort":"high"}` hoặc `{"max_tokens":N}` (không gửi cả hai). Luồng trả `choices[].delta.reasoning_details` (các khối `reasoning.text` kèm `signature`, `reasoning.encrypted`, `reasoning.summary`). Khi tiếp tục vòng lặp công cụ phải gửi lại `reasoning_details` của tin nhắn trợ lý trước đó, nguyên vẹn, đúng thứ tự. Gửi `reasoning` cho model không suy nghĩ: chưa xác minh.

## OpenAI (`/v1/chat/completions`)

- `reasoning_effort`: `none|minimal|low|medium|high|xhigh|max`, mỗi model nhận một tập con; gửi giá trị model không nhận thì 400.
- PDF: phần `{"type":"file","file":{"filename":"x.pdf","file_data":"data:application/pdf;base64,…"}}`, mỗi tệp dưới 50 MB; cần model nhận ảnh.

## Quyết định cho bản 4.1

| Nguồn | PDF quét | Suy nghĩ kỹ |
|---|---|---|
| Anthropic | Gửi nguyên tệp (khối `document`) | Bật: `adaptive` + `effort` cao hơn mặc định (model đời mới); `enabled` + `budget_tokens` (model cũ). Luôn giữ và gửi lại khối suy nghĩ; kèm `drop_block` để máy chủ tự bỏ khối hết hợp lệ; gặp 400 về khối suy nghĩ thì bỏ khối và gửi lại một lần |
| Gemini | Gửi nguyên tệp (`inlineData`) | Bật: `thinkingLevel: "high"` (3.x, 4) hoặc `thinkingBudget: -1` → mức cao (2.5); gặp 400 thì gửi lại không có `thinkingConfig` |
| OpenRouter | Gửi nguyên tệp, chỉ rõ engine `native` (không tự chuyển sang OCR có phí); model không đọc được tệp thì báo rõ | Bật: `reasoning: {"effort":"high"}`; giữ và gửi lại `reasoning_details` |
| OpenAI | Gửi nguyên tệp (phần `file`) | Bật: `reasoning_effort: "high"`; gặp 400 thì gửi lại không có tham số |
| Nguồn khác | Báo rõ không đọc được PDF quét, gợi ý đổi sang Anthropic hoặc Gemini | Công tắc không có tác dụng, khung ghi rõ |
