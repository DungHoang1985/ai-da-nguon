# So song song với add-in Claude chính thức; PDF quét; Suy nghĩ kỹ (16/11/2026)

Việc nợ từ Đợt 1A (tiêu chí 8) và Đợt 1B. Anh ngồi máy; add-in Claude chính thức chạy bằng tài khoản của anh (Sonnet 5.5, mức High, chế độ "Ask before edits"). AI Đa Nguồn: bản đã phát hành 4.4.0, trừ các dòng ghi "bản sửa" (bản cục bộ nhánh `claude/so-song-song`, nạp bằng `cdp.mjs serve`).

Cách chạy: hợp đồng thử 54 trang (`kiem-thu/hop-dong-thu.mjs`: 165 đoạn, 30 sửa đổi và 5 nhận xét của "Đối tác"), mỗi model một bản mới. Cả hai khung nằm trong cùng một Word, điều khiển bằng mã qua cổng gỡ lỗi (khung chính thức là trang `pivot.claude.ai`); nút Claude trên thanh công cụ bấm bằng giao diện trợ năng của Windows. Chấm bài d, e bằng mã đọc lại văn bản (số sửa đổi, câu trả lời trong từng nhận xét, số và kiểu của khoản mới). Bằng chứng: `tasks/so-song-song-bang-chung/`.

Năm bài: (a) "Điều 18.2(b) nói gì?" · (b) "Lãi chậm thanh toán được quy định ở điều nào và mức bao nhiêu?" · (c) tóm tắt đủ 30 sửa đổi của đối tác theo số điều · (d) xử lý cả 5 nhận xét bằng Track Changes và trả lời từng nhận xét · (e) thêm một khoản ngay sau Điều 8.2.

## Bảng đối chiếu 5 bài

| Bên chạy | a | b | c | d | e | Bằng chứng |
|---|---|---|---|---|---|---|
| Claude chính thức · Sonnet 5.5 | Đạt, 11 giây | Đạt, 14 giây | Đạt, đủ 30, 33 giây | Lượt một: sửa 5/5, chưa trả lời nhận xét (chờ bấm Apply). Chạy lại hai lượt: sửa 5/5, trả lời 5/5, không đánh dấu đã xử lý; 36 + 20 giây | Đạt (8.3 mới, 8.4 cũ, cùng kiểu), 59 giây | `claude-sonnet55-*` (bài a chỉ có trong nhật ký phiên) |
| AI Đa Nguồn · Sonnet 5.5 | Đạt, 15 giây | Đạt, 9 giây | Đạt, đủ 30, 26 giây | Sửa 5/5, trả lời 5/5 trong một lượt, 27 giây | Đạt, 11 giây | `minh-sonnet55` |
| AI Đa Nguồn · Haiku 5.5 | Đạt, 6 giây | Đạt, 5 giây | Đạt, 23 giây | Sửa 5/5, trả lời 5/5, 23 giây | Đạt, 9 giây | `minh-haiku55` |
| AI Đa Nguồn · Cerebras gpt-oss-120b (miễn phí) | Đạt | Đạt | Đạt, đủ 30 | Không đạt: hết hạn mức của Cerebras giữa chừng, chưa sửa gì | Đạt | `minh-cerebras-oss120b` |
| AI Đa Nguồn · Cohere command-a-plus (miễn phí) | Trích đúng, diễn giải sai ý | Đạt | Đạt, đủ 30 | Sửa 3/5 rồi nhà cung cấp báo lỗi, không trả lời nhận xét | Đạt | `minh-cohere-aplus` |
| AI Đa Nguồn · Groq gpt-oss-120b (miễn phí) | Không chạy được: yêu cầu lớn hơn 8.000 token mỗi phút của gói miễn phí (câu báo tiếng Việt, đúng) | | | | | `minh-groq-oss120b` |
| AI Đa Nguồn · OpenAI (gpt-6.1-sol, gpt-5.4-mini) | Chưa kiểm được: OpenAI trả lỗi mà trình duyệt không cho đọc (danh sách model vẫn tải được nên key còn nhận; nhiều khả năng do số dư hoặc quyền dùng model) | | | | | `minh-gpt61sol`, `minh-gpt54mini` |
| AI Đa Nguồn · DeepSeek | Chưa kiểm được: tài khoản hết tiền | | | | | `minh-deepseekflash` |
| AI Đa Nguồn · Gemini 3.8 Flash | Bản 4.4.0: **lỗi 400 ở mọi yêu cầu** (lỗi của add-in, xem dưới). Bản sửa: trả lời đúng bài a một lần (3 lần gọi); chạy đủ 5 bài thì Google trả 503, im lặng rồi báo hết hạn mức trong ngày → chưa kiểm được đủ | | | | | `minh-gemini38flash`, `minh-gemini38flash-ban-sua` |

Nhận xét:
- Cùng model Sonnet 5.5, AI Đa Nguồn **đạt ngang** bản chính thức ở a, b, c, e và **gọn hơn ở d** (một lượt, không phải bấm Apply từng thẻ). Bản chính thức có thẻ đề xuất kèm lý do cho từng chỗ sửa và liên kết bấm tới điều khoản — AI Đa Nguồn chưa có thẻ kiểu đó.
- Bài b: hai bên nói khác nhau về Điều 6.4 (bản chính thức: 45 ngày — chữ gốc; AI Đa Nguồn: 90 ngày — chữ đối tác đang sửa). **Không bên nào nói rõ đây là chỗ đang chờ duyệt.** Nên sửa: khi dẫn một đoạn đang có sửa đổi chờ duyệt, AI Đa Nguồn phải nêu cả hai giá trị.
- Model miễn phí đọc và tóm tắt tốt, nhưng việc sửa nhiều bước (bài d) hay gãy vì hạn mức hoặc lỗi nhà cung cấp.

## PDF quét 3 trang (ảnh, không có lớp chữ; ba số liệu phải đúng)

| Bên chạy | Kết quả | Bằng chứng |
|---|---|---|
| AI Đa Nguồn · Sonnet 5.5 | Đạt 3/3, 8 giây | `pdf-sonnet55` |
| AI Đa Nguồn · Haiku 5.5 | Đạt 3/3, 7 giây | `pdf-haiku55` |
| Claude chính thức · Sonnet 5.5 | Đạt 3/3, 20 giây; đọc tệp **bằng công cụ, theo từng trang**, dẫn trang cho từng số và tự kiểm chéo phép tính (5% của giá trị hợp đồng, tổng ba khoản) | `pdf-claude-chinh-thuc`, `pdf-claude-chinh-thuc-luot2` |
| AI Đa Nguồn · Cerebras, Cohere | Báo rõ không đọc được PDF quét, tệp không gửi đi, gợi ý đổi nguồn (đúng thiết kế) | `pdf-cerebras`, `pdf-cohere` |
| AI Đa Nguồn · Gemini 3.8 Flash | **Chưa kiểm được** — Google báo hết hạn mức trong ngày | `pdf-gemini38flash` |
| OpenAI, OpenRouter (trả tiền) | **Chưa kiểm được** — OpenAI lỗi như trên; OpenRouter số dư bằng không, lượt miễn phí để lại theo ý anh | — |

## Suy nghĩ kỹ

- Gemini 3.8 Flash: yêu cầu gửi đi có `thinkingConfig.thinkingLevel = high` (đúng tham số), nhưng cả 4 lần Google trả 429 hết hạn mức → **Chưa kiểm được** (`nghi-ky-gemini38flash`).
- OpenAI, OpenRouter: **Chưa kiểm được**, cùng lý do ở trên. Anthropic đã đạt ở Đợt 1B.

## Hai lỗi của AI Đa Nguồn tìm ra và đã sửa (chưa phát hành)

1. **Mọi model Gemini bị từ chối 400 trong Word từ bản 4.2.1.** Công cụ `word_create_admin_document` khai `font_size` có `enum: [13, 14]` kiểu số; Gemini chỉ nhận enum chuỗi nên từ chối cả yêu cầu. Sửa ở `toGeminiSchema` (`src/core/agent/llm.ts`): enum không phải chuỗi thì bỏ, ghi giá trị cho phép vào mô tả. Có kiểm thử máy. Chạy thật trên bản sửa: hết lỗi 400, Gemini gọi công cụ và trả lời đúng.
2. **Khung treo vô hạn khi nhà cung cấp nhận yêu cầu rồi im lặng** (gặp với Gemini 3.7 và 3.8 Flash sau lỗi 503; trước đây chỉ thoát bằng nút Dừng). Thêm hạn chờ `STALL_MS` = 150 giây không có byte nào thì hủy và báo "không nhận được phản hồi trong 150 giây…", không tự gửi lại. 3 kiểm thử máy; Chạy thật: khung hiện đúng câu báo sau 152 giây.

`npm test` 474/474, `npx tsc --noEmit -p .` sạch, `npm run build` đạt. Commit cục bộ trên nhánh `claude/so-song-song`.

## Kinh nghiệm xử lý tài liệu quét từ ba dự án khác của anh (khảo sát chỉ đọc)

Nguồn: `D:\he-thong-phap-ly-bds` (văn bản pháp luật tiếng Việt), `D:\ky-mon-don-giap` và `D:\nhat-trach-dai-quai` (sách chữ Hán quét). Điểm chung: không dự án nào gửi cả tệp cho model; tất cả tách theo trang, dựng ảnh từng trang rồi đối chiếu bằng máy. Bản chính thức cũng đọc theo từng trang bằng công cụ.

Những điều đã trả giá ở các dự án đó:
- OCR máy làm mất dấu tiếng Việt; model đọc ảnh đạt từ 99% từ khớp, nhưng vẫn có ca sai một con số (30 thành 80 ngày), điểm a), b) thành 1., 2.
- Model tự "sửa" lỗi in cho hợp lý, chép sót hư từ, suy theo quy luật thay vì đọc; bị cắt giữa chừng khi hết lượt đầu ra; thiếu trang; mất nửa câu ở chỗ vắt trang.
- Cách chặn: lời dặn chép nguyên văn ("đúng chữ in kể cả khi sai, không tóm lược, chữ không đọc được ghi [?], không đoán; số và ngày chép đúng kiểu in; bảng giữ hàng cột"); dấu kết trang; đếm trang vào bằng trang ra; máy kiểm dãy Điều – khoản – điểm; đọc hai lần độc lập cho số, ngày, số hiệu rồi máy so, lệch thì đưa người xem; bản đọc từ ảnh "chỉ để đối chiếu, không làm căn cứ trích dẫn".
- Phân loại trang: ảnh phủ từ 0,8 trang và dưới 300 ký tự lớp chữ là trang quét (ngưỡng 300 vì dòng chữ ký số); ảnh 150–200 dpi là đủ, chỗ khó phóng 300 dpi; khoảng 1.800 token một ảnh trang.

Đề xuất cho add-in (chưa làm, cần anh duyệt vì là một đợt việc riêng):
1. Công cụ đọc tệp đính kèm **theo trang** (`read_attachment_pages`): phân loại từng trang bằng pdf.js; trang có lớp chữ lấy chữ; trang quét dựng ảnh và gửi cho model đọc theo yêu cầu, kèm số trang — thay cho gửi cả tệp vào mọi lượt (hiện tệp nằm lại trong lịch sử và tốn token mỗi lượt).
2. Lời dặn chép nguyên văn và dẫn trang cho mọi số liệu lấy từ bản quét; câu trả lời ghi rõ "đọc từ bản quét, cần đối chiếu bản gốc".
3. Máy kiểm sau khi đọc: đếm trang, tỉ lệ chữ có dấu, rút số / ngày / số hiệu; tùy chọn đọc lần hai độc lập cho các con số then chốt và báo chỗ lệch.
4. Cảnh báo bảo mật khi gửi tài liệu thật cho nguồn miễn phí hoặc qua OpenRouter.

## Dọn dẹp

Hội thoại thử trong khung AI Đa Nguồn đã xóa theo tiêu đề; nguồn trả về Anthropic · Claude Opus 5.5; công tắc Suy nghĩ kỹ tắt. Các hội thoại thử trong add-in Claude chính thức (5 bài, 2 lượt PDF) còn trong lịch sử tài khoản của anh — tôi không xóa. Tệp thử nằm ở thư mục tạm của phiên, không vào kho.
