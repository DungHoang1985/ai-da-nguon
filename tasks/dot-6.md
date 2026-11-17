# Đợt 6 — thêm nhà cung cấp nói cùng "tiếng" với Anthropic (kế hoạch, chờ anh duyệt)

Anh nêu ngày 17/11/2026: add-in dựng theo bản mẫu của Claude, nên bổ sung các nguồn nhận đúng kiểu yêu cầu của Anthropic (Messages API: `POST …/v1/messages`).

Trạng thái: **mới là kế hoạch, chưa sửa mã.** Phần thăm dò dưới đây làm ngày 17/11/2026, không dùng key nào.

## Vì sao đáng làm

Add-in đã có một đường gọi kiểu Anthropic hoàn chỉnh (`llm.ts`, nhánh `kind: "anthropic"`): bộ nhớ đệm lời nhắc, khối suy nghĩ gửi lại nguyên văn, PDF gửi nguyên tệp, ảnh, công cụ. Hiện chỉ một nguồn dùng đường này (`api.anthropic.com`). Mười nguồn còn lại đi đường OpenAI hoặc Gemini, nên cùng model Claude mà gọi qua OpenRouter thì mất bộ nhớ đệm chủ động và PDF nguyên tệp.

Nguồn nào nói cùng tiếng thì dùng lại được đường này, chỉ khác địa chỉ, cách gửi key và vài giới hạn.

## Hiện trạng đã kiểm

Hai phép đo cho mỗi địa chỉ, giả làm trình duyệt (gửi kèm `Origin`): (1) yêu cầu dò trước (`OPTIONS`) có được trả tiêu đề cho phép không; (2) gọi thật không kèm key có trả lỗi 401 đọc được không. Chưa đo trong khung add-in thật.

### Nhóm 1 — chính model Claude, qua cửa khác

| Nguồn | Địa chỉ | Cách gửi key | Add-in gọi được? | Ghi chú từ tài liệu chính thức |
|---|---|---|---|---|
| OpenRouter (kênh Messages) | `https://openrouter.ai/api/v1/messages` | `Authorization: Bearer` | Được (đo: cho phép mọi nguồn gốc, nhận `Authorization` và `X-Api-Key`) | Nhận chữ, ảnh, PDF, công cụ, suy nghĩ, `cache_control`. Anh đã có key và đăng nhập một lần bấm. Tài liệu không nói rõ mọi tính năng chạy với model không phải Claude |
| Amazon Bedrock | `https://bedrock-mantle.<vùng>.api.aws/anthropic/v1/messages` | `x-api-key` (khóa Bedrock) | Được ở bước dò trước (đo vùng us-east-1: cho phép mọi nguồn gốc, mọi tiêu đề); chưa đo gọi thật | Cần tài khoản AWS và khóa Bedrock; phải chọn vùng |
| Microsoft Foundry | `https://<tên tài nguyên>.services.ai.azure.com/anthropic/v1/messages` | `api-key` hoặc `Authorization: Bearer` | Chưa đo được (địa chỉ theo tài nguyên của từng người) | Tên model là tên bản triển khai; không có API liệt kê model |
| Vercel AI Gateway | `https://ai-gateway.vercel.sh/v1/messages` | `x-api-key` hoặc `Authorization: Bearer` | Được ở bước dò trước; lỗi 401 không kèm tiêu đề cho phép nên key sai sẽ hiện thành "không kết nối được" | Tên model dạng `anthropic/claude-…`; `cache_control` được chuyển tiếp cho model Claude |

### Nhóm 2 — model khác, nhận kiểu yêu cầu của Anthropic

| Nguồn | Địa chỉ | Cách gửi key | Add-in gọi được? | Giới hạn theo tài liệu chính thức |
|---|---|---|---|---|
| DeepSeek | `https://api.deepseek.com/anthropic/v1/messages` | `x-api-key` hoặc `Authorization: Bearer` | Được (đo: nhận đủ `x-api-key`, `anthropic-version`, `anthropic-beta`; lỗi 401 đọc được) | Model `deepseek-flash`, `deepseek-v4-pro`. Bỏ qua `cache_control`. PDF base64 bị từ chối 400. Ảnh: chỉ `deepseek-flash`. Suy nghĩ: `adaptive` coi như bật, `budget_tokens` bị bỏ qua; khi đang suy nghĩ thì `tool_choice` ép gọi công cụ bị 400. Không nhận khối `redacted_thinking` |
| MiniMax | `https://api.minimax.io/anthropic/v1/messages` | `Authorization: Bearer` | Được nếu CHỈ gửi `Authorization` (đo: bước dò trước không cho `x-api-key`, `anthropic-version`) | Model `MiniMax-M3`, `MiniMax-M3.1-Flash`. Nhận ảnh. `tool_choice` chỉ nhận `auto`. M3.1-Flash luôn suy nghĩ, gửi `disabled` bị 400. Phải giữ khối suy nghĩ trong lịch sử. Tài liệu không nói về PDF |
| Kimi (Moonshot) | `https://api.moonshot.ai/anthropic/v1/messages` | `Authorization: Bearer` | Được nếu CHỈ gửi `Authorization` và `Content-Type` (đo: xin thêm tiêu đề `anthropic-*` thì bước dò trước không trả phép) | Model `kimi-k3.1-pro`. Trang hướng dẫn chỉ nói "tương thích Messages API", chưa đọc được trang chi tiết |
| Z.ai (GLM) | `https://api.z.ai/api/anthropic/v1/messages` | khóa Z.ai | **Không** (đo hai lần: bước dò trước không trả tiêu đề cho phép nào) | Tài liệu viết cho gói GLM Coding Plan; không nói key trả theo lượt có dùng được không |
| xAI (Grok) | `https://api.x.ai/v1/messages` | — | Được về mặt kỹ thuật | **Loại**: tài liệu xAI ghi kênh tương thích Anthropic đã ngừng hỗ trợ |
| Alibaba Qwen | chưa tra được địa chỉ chính thức | — | Chưa đo | Địa chỉ tôi thử trả 403; phải đọc tài liệu Model Studio trước |

Đã thử và loại: Groq, Fireworks (không có kênh Messages — trả 404), Ollama Cloud (405).

## Thiết kế

1. **Tách "phương ngữ" khỏi nhà cung cấp.** Thêm vào `Provider` một cấu hình `anthropicDialect` cho nguồn `kind: "anthropic"`:
   - cách gửi key: `x-api-key` · `bearer` · `api-key`;
   - có gửi `anthropic-version`, `anthropic-beta`, `anthropic-dangerous-direct-browser-access` không (MiniMax, Kimi: không gửi);
   - có `cache_control` không; có nhận PDF nguyên tệp không; ảnh theo model;
   - suy nghĩ: như Anthropic · luôn bật · chỉ `enabled` · không có;
   - có ép `tool_choice` được không;
   - danh sách model: lấy từ API hay danh sách cố định ghi theo tài liệu.
   `api.anthropic.com` giữ nguyên hành vi hiện tại (mặc định của cấu hình).
2. **Một nguồn, hai kênh** cho OpenRouter và DeepSeek (đang đi đường OpenAI): model Claude trên OpenRouter đi kênh Messages; còn lại giữ đường cũ. Không đổi key, không đổi đăng nhập.
3. **"Nguồn tùy chỉnh kiểu Anthropic"**: anh nhập địa chỉ gốc, key, tên model. Dùng cho Foundry, cổng trung chuyển, nguồn mới ra sau này. Chỉ nhận `https`, hiện rõ tên miền sẽ nhận key, không tự điền.
4. Không đổi lời nhắc hệ thống, bộ công cụ, lịch sử chỉ thêm. Lỗi của nguồn mới vẫn qua `maskSecrets` và bảng phân loại lỗi key / tiền.

## Các bước

| Bước | Việc | Cần anh? |
|---|---|---|
| 6.0 | Đo lại trong khung add-in thật (desktop ẩn): gọi từng địa chỉ không kèm key, ghi bảng. Đọc nốt tài liệu Kimi, Qwen, Bedrock | Không |
| 6.1 | Cấu hình phương ngữ + kiểm thử máy cho từng tổ hợp tiêu đề, suy nghĩ, PDF, ảnh | Không |
| 6.2 | OpenRouter kênh Messages cho model Claude; DeepSeek kênh Anthropic | Không (đã có key OpenRouter; DeepSeek cần key nếu anh chưa có) |
| 6.3 | Thêm nguồn anh chọn ở nhóm còn lại (MiniMax, Kimi, Bedrock, Vercel) | Có: tạo key ở từng nguồn |
| 6.4 | Nguồn tùy chỉnh kiểu Anthropic (gồm Foundry) | Có nếu anh muốn thử với tài nguyên thật |
| 6.5 | Tự kiểm tra ba ứng dụng, cập nhật bảng so sánh và HANDOFF, xin duyệt phát hành | Duyệt phát hành |

## Tiêu chí nghiệm thu

| # | Tiêu chí | Mức bằng chứng |
|---|---|---|
| 1 | Mỗi nguồn mới: tiêu đề gửi đi đúng bảng phương ngữ; nguồn Anthropic gốc không đổi một byte nào trong yêu cầu | Kiểm thử máy |
| 2 | Mỗi nguồn có key: một bài Word đọc rồi sửa có Track Changes, có gọi công cụ, ghi số lần gọi AI | Chạy thật |
| 3 | Claude qua OpenRouter kênh Messages: lượt thứ hai có "đọc từ bộ nhớ đệm" lớn hơn 0; PDF quét 3 trang đúng 3/3; Suy nghĩ kỹ chạy | Chạy thật |
| 4 | Nguồn không nhận PDF hoặc ảnh: báo rõ, không gửi, không bịa | Chạy thật |
| 5 | Key sai ở mỗi nguồn: câu lỗi tiếng Việt đúng loại, không lộ key | Chạy thật |
| 6 | Nguồn bị chặn gọi từ add-in (Z.ai): hiện lý do ngay khi chọn, như SambaNova và NVIDIA | Tự kiểm bằng mã trong khung |
| 7 | Tự kiểm tra Word, Excel, PowerPoint không thụt lùi | Tự kiểm tra |

## Việc buộc phải có anh ở máy

- Tạo key (và nạp tiền nếu cần) ở từng nguồn anh chọn; dán key vào Cài đặt của add-in. Phiên không tạo key, không đăng nhập, không bấm đồng ý điều khoản thay anh.
- Duyệt phát hành.

## Rủi ro đã thấy

- Tài liệu của các nguồn nhóm 2 đổi nhanh; tên model trong bảng là theo tài liệu đọc ngày 17/11/2026.
- Phép đo bằng `curl` chưa chắc trùng với khung add-in thật: bước 6.0 đo lại trước khi viết mã.
- Nguồn nhóm 2 bỏ qua `cache_control`: không có bộ nhớ đệm chủ động, mỗi lượt tính đủ token.
- Vercel: key sai sẽ hiện thành lỗi kết nối chứ không phải "key sai", trừ khi tìm được cách phân biệt.

## Nguồn tài liệu đã đọc

- https://api-docs.deepseek.com/guides/anthropic_api
- https://platform.minimax.io/docs/api-reference/text-anthropic-api
- https://platform.kimi.ai/docs/guide/agent-support
- https://docs.z.ai/scenario-example/develop-tools/claude
- https://openrouter.ai/docs/api/api-reference/anthropic-messages/create-a-message.md
- https://vercel.com/docs/ai-gateway/sdks-and-apis/anthropic-messages-api
- https://platform.claude.com/docs/en/build-with-claude/claude-in-amazon-bedrock
- https://platform.claude.com/docs/en/build-with-claude/claude-in-microsoft-foundry
- https://docs.x.ai/developers/rest-api-reference/inference/legacy.md
