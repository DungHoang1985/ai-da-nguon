/** Đọc tệp PDF đính kèm THEO TRANG (Đợt 5): chọn trang, phân loại trang quét / trang có lớp chữ, đường đi của PDF quét lúc gửi,
 *  lời dặn chép nguyên văn, cảnh báo riêng tư. Phần này không đụng pdf.js hay DOM nên kiểm thử được bằng npm test;
 *  phần dựng ảnh nằm ở pdfRender.ts. */

/** Trang có ít hơn ngần này ký tự lớp chữ thì coi là trang quét (ngưỡng 300: trang quét có chữ ký số vẫn mang một dòng chữ) */
export const SCAN_TEXT_MIN_CHARS = 300;
/** Vòng lặp tác tử chỉ chuyển tối đa 4 ảnh cho một lần gọi công cụ */
export const MAX_SCAN_PAGES_PER_CALL = 4;
export const MAX_TEXT_PAGES_PER_CALL = 20;
/** Chữ tối đa của một lần gọi (các trang có lớp chữ). Phải nằm DƯỚI trần kết quả công cụ của vòng lặp (30.000 ký tự sau khi
 *  đóng gói JSON): vượt thì kết quả bị cắt đuôi, mất lời dặn và phần chữ cuối mà model không biết (soát mã 18/11/2026). */
export const MAX_PAGE_TEXT_CHARS = 20000;
/** PDF quét dài hơn ngần này trang thì không gửi nguyên tệp nữa (tệp nằm lại trong lịch sử, tốn token mọi lượt): AI đọc theo trang */
export const WHOLE_PDF_MAX_PAGES = 5;

export function isScanPage(textChars: number): boolean {
  return textChars < SCAN_TEXT_MIN_CHARS;
}

/** "1-3, 5" → [1,2,3,5]. Bỏ trùng, giữ thứ tự tăng. Sai thì trả error (câu tiếng Việt để model sửa lời gọi). */
export function parsePageSpec(spec: unknown, total: number): { pages: number[]; error?: string } {
  const s = String(spec ?? "").trim();
  if (!s) return { pages: [], error: "Thiếu pages. Ví dụ: \"1-4\" hoặc \"2, 5, 7\"." };
  const set = new Set<number>();
  // "1 - 3" (có dấu cách quanh gạch nối) là một khoảng
  for (const part of s.replace(/\s*[-–]\s*/g, "-").split(/[,;\s]+/).filter(Boolean)) {
    const m = /^(\d+)(?:\s*[-–]\s*(\d+))?$/.exec(part);
    if (!m) return { pages: [], error: `Không hiểu "${part}" trong pages. Ví dụ: "1-4" hoặc "2, 5, 7".` };
    const a = Number(m[1]);
    const b = m[2] ? Number(m[2]) : a;
    if (a < 1 || b < a) return { pages: [], error: `Khoảng trang "${part}" không hợp lệ.` };
    if (a > total) return { pages: [], error: `Tệp chỉ có ${total} trang (yêu cầu trang ${a}).` };
    for (let i = a; i <= Math.min(b, total); i++) set.add(i);
    if (set.size > 400) break;
  }
  return { pages: [...set].sort((x, y) => x - y) };
}

/** Rút gọn dãy trang để ghi vào kết quả: [1,2,3,5] → "1-3, 5" */
export function formatPages(pages: number[]): string {
  const out: string[] = [];
  const p = [...pages].sort((a, b) => a - b);
  for (let i = 0; i < p.length; i++) {
    let j = i;
    while (j + 1 < p.length && p[j + 1] === p[j] + 1) j++;
    out.push(j > i ? `${p[i]}-${p[j]}` : String(p[i]));
    i = j;
  }
  return out.join(", ");
}

/** Chia các trang được yêu cầu thành phần đưa ngay và phần để lần gọi sau: tối đa 4 trang quét và 20 trang chữ mỗi lần.
 *  Dừng ở trang đầu tiên vượt trần để trang đưa ra luôn liền mạch theo thứ tự (không nhảy cóc qua trang quét). */
export function planPages(wanted: number[], isScan: (page: number) => boolean): { now: number[]; later: number[] } {
  const now: number[] = [];
  let scans = 0;
  let texts = 0;
  let i = 0;
  for (; i < wanted.length; i++) {
    const scan = isScan(wanted[i]);
    if (scan ? scans >= MAX_SCAN_PAGES_PER_CALL : texts >= MAX_TEXT_PAGES_PER_CALL) break;
    if (scan) scans++;
    else texts++;
    now.push(wanted[i]);
  }
  return { now, later: wanted.slice(i) };
}

/** Lời dặn đi kèm mỗi lần đưa ảnh trang quét (nằm trong kết quả công cụ, không nằm trong lời nhắc hệ thống) */
/** Ghi vào kết quả công cụ khi model không (chắc) xem được ảnh: không gửi ảnh trang quét, không để model đoán */
export const NO_VISION_NOTE =
  "Các trang quét KHÔNG được gửi: model đang dùng không nhận ảnh (hoặc danh mục không ghi rõ là nhận ảnh). Không được đoán nội dung các trang này; hãy nói rõ với người dùng và gợi ý chọn một model nhận ảnh (ví dụ nguồn Anthropic hoặc Gemini) rồi hỏi lại.";

export const SCAN_RULES =
  "Các trang quét được gửi dưới dạng ảnh ở tin nhắn ngay sau kết quả này. Khi đọc: chép ĐÚNG chữ in trên ảnh, kể cả khi chữ in có vẻ sai — không tự sửa cho hợp lý, không suy theo quy luật; " +
  "chữ, số không đọc được thì ghi [?], không đoán; số tiền, ngày tháng, số hiệu chép đúng kiểu in; bảng giữ đúng hàng cột; ký hiệu điểm a), b) giữ nguyên, không đổi thành 1., 2. " +
  "Mọi số liệu lấy từ trang quét phải ghi kèm số trang, và nói với người dùng một lần rằng nội dung đọc từ bản quét, cần đối chiếu bản gốc trước khi dùng làm căn cứ. " +
  "Chưa đọc trang nào thì không được nói về trang đó: còn trang cần cho câu hỏi thì gọi tiếp read_attachment_pages.";

export type ScanRoute = "whole" | "pages" | "none";

/** Model có chắc xem được ảnh trang không. vision = điều danh mục biết về model (undefined = không rõ).
 *  Nguồn đọc được nguyên tệp PDF thì model không rõ vẫn coi là xem được (như trước); nguồn khác phải được ghi rõ là nhận ảnh —
 *  gửi ảnh cho model chỉ đọc chữ là bị từ chối cả lượt (Cerebras: "Content type 'image_url' is not supported", đo 17/11/2026). */
export function canSeePages(o: { pdfCapable: boolean; vision?: boolean }): boolean {
  return o.vision === true || (o.vision === undefined && o.pdfCapable);
}

/** PDF quét lúc gửi đi đường nào: nguyên tệp (như trước) · AI đọc theo trang bằng công cụ · không đọc được.
 *  tools = lượt này model có công cụ (bật "AI tự thao tác" và model gọi được công cụ). */
export function routeScanPdf(o: { pages: number; pdfCapable: boolean; vision?: boolean; tools: boolean }): ScanRoute {
  if (o.vision === false) return "none";
  if (o.pdfCapable) return o.pages <= WHOLE_PDF_MAX_PAGES || !o.tools ? "whole" : "pages";
  return o.tools && canSeePages(o) ? "pages" : "none";
}

/** Dòng nhắc riêng tư khi tệp đính kèm sắp được gửi cho nguồn miễn phí hoặc qua nơi trung chuyển. Không phải thì trả "". */
/** Nguồn có gói miễn phí phổ biến, key không cho biết thuộc gói nào */
const FREE_TIER_PROVIDERS = new Set(["gemini", "groq", "cerebras", "mistral"]);

export function attachmentPrivacyNote(o: { providerId: string; providerName: string; modelId: string; free?: boolean }): string {
  const freeModel = /:free$/i.test(o.modelId) || !!o.free;
  // Mọi kênh đi qua OpenRouter (kể cả kênh Claude kiểu Anthropic) đều là nơi trung chuyển
  if (/^openrouter(-|$)/.test(o.providerId)) {
    return freeModel
      ? `Tệp đính kèm được gửi qua OpenRouter tới một model miễn phí: nhà cung cấp model miễn phí thường được phép lưu và dùng nội dung để huấn luyện. Tài liệu thật, có thông tin mật thì nên dùng nguồn trả tiền.`
      : `Tệp đính kèm được gửi qua OpenRouter (nơi trung chuyển) rồi tới nhà cung cấp model. Tài liệu có thông tin mật thì cân nhắc gửi thẳng tới nhà cung cấp.`;
  }
  if (freeModel) return `Tệp đính kèm được gửi tới ${o.providerName} theo gói miễn phí: gói miễn phí thường cho phép nhà cung cấp lưu và dùng nội dung để cải thiện model. Tài liệu thật, có thông tin mật thì nên dùng nguồn trả tiền.`;
  // Nguồn có gói miễn phí dùng chung một loại key với gói trả tiền: add-in không biết key thuộc gói nào → nhắc có điều kiện
  // (soát mã Đợt 5, mục 4)
  if (FREE_TIER_PROVIDERS.has(o.providerId))
    return `Nếu key ${o.providerName} của bạn thuộc gói miễn phí: điều khoản gói miễn phí có thể cho phép nhà cung cấp lưu và dùng nội dung tệp để cải thiện model (xem điều khoản của nhà cung cấp). Tài liệu có thông tin mật thì nên dùng gói trả tiền.`;
  return "";
}
