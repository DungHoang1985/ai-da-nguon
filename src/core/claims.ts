/** Nhận diện ca "AI nói đã sửa nhưng lượt đó không có thao tác ghi nào" (Đợt 3, việc A2) — hàm thuần, kiểm thử được.
 *  Kết quả chỉ dùng để HIỆN cảnh báo trên khung; không đi vào lịch sử gửi model, không đổi lời nhắc hệ thống.
 *  Bối cảnh: 16/11/2026 một model miễn phí trả lời "Đã sửa…" sau 3 giây mà không gọi công cụ nào, văn bản không đổi. */

/** Động từ chỉ việc thay đổi tài liệu (không gồm việc đọc, kiểm, tìm, tóm tắt) */
const VERBS =
  "sửa|chỉnh sửa|chỉnh|thay thế|thay|thêm|chèn|xóa|xoá|ghi|cập nhật|đổi|điền|soạn|tạo|dựng|định dạng|viết lại|viết|bổ sung|áp dụng|chuyển|di chuyển|đặt lại|đặt|tô|gỡ|bỏ|loại bỏ|đánh dấu|chấp nhận|từ chối|trả lời nhận xét|sắp xếp|gộp|tách|nhân bản|lưu|hoàn tất việc sửa|điều chỉnh|làm lại|trình bày lại|căn chỉnh|căn|vẽ";
const VERB = `(?:${VERBS})`;
/** Chủ ngữ là chính AI */
const FIRST = "(?:tôi|em|mình|tớ|trợ lý|ai)";
/** Chủ ngữ là một phần của tài liệu (câu bị động: "Đoạn 3 đã được sửa") */
const OBJECT =
  "(?:văn bản|tài liệu|hợp đồng|công văn|tờ trình|đoạn|điều|khoản|điểm|mục|câu|ô|bảng|sheet|trang tính|sổ tính|cột|hàng|dòng|công thức|biểu đồ|slide|trang|bài trình chiếu|tiêu đề|nội dung|số liệu|dữ liệu|tệp|file|nhận xét|sửa đổi|thay đổi|định dạng|lề|phông|chữ)";

const L = "[^\\p{L}\\p{N}]";
const reFirst = new RegExp(`(?:^|${L})${FIRST}\\s+(?:cũng\\s+|vừa\\s+)?(?:đã|vừa)\\s+(?:được\\s+)?(?:tiến hành\\s+|thực hiện\\s+(?:việc\\s+)?)?${VERB}(?=$|${L})`, "iu");
const reInitial = new RegExp(`^(?:đã|vừa)\\s+(?:tiến hành\\s+|thực hiện\\s+(?:việc\\s+)?)?${VERB}(?=$|${L})`, "iu");
const rePassive = new RegExp(`(?:^|${L})${OBJECT}(?:[^.;:!?\\n]|\\.(?=\\d)){0,60}?\\s(?:đã|vừa)\\s+được\\s+(?:tôi\\s+)?${VERB}(?=$|${L})`, "iu");
const reEnglish = /\b(?:i(?:'ve| have)|i)\s+(?:now\s+|just\s+|already\s+)?(?:updated|changed|replaced|inserted|added|deleted|removed|edited|fixed|modified|applied|rewritten|rewrote|created|formatted|corrected)\b|\b(?:has|have) been (?:updated|changed|replaced|inserted|added|deleted|removed|edited|fixed|modified|applied|rewritten|created|formatted|corrected)\b/i;
/** Câu không phải lời khẳng định đã làm: phủ định, tương lai, điều kiện, đề xuất, hướng dẫn người dùng tự làm */
const reHedge = new RegExp(
  `(?:^|${L})(?:chưa|không thể|không\\s+\\p{L}+\\s+được|chẳng|sẽ|nếu|giả sử|khi nào|sau khi|một khi|trước khi|có thể|nên|cần|hãy|vui lòng|xin hãy|đề xuất|đề nghị|gợi ý|khuyến nghị|dự kiến|muốn|bạn đã|anh đã|chị đã|người dùng đã|đối tác|ai đó|tác giả)(?=$|${L})`,
  "iu"
);
const reIntent = new RegExp(`(?:^|${L})(?:${VERBS}|làm|hoàn thiện|rút gọn|dịch|chuẩn hóa|chuẩn hoá|xử lý|fix|edit|update|rewrite|replace|insert|add|delete|remove|format|write|create|change)(?=$|${L})`, "iu");

export type ClaimStrength = "none" | "medium" | "strong";

export interface ClaimInfo {
  strength: ClaimStrength;
  /** Câu khẳng định tìm được (rút gọn) */
  sentence?: string;
}

/** Bỏ phần không phải lời của AI về việc mình làm: khối mã, phần trích trong ngoặc kép, dòng trích dẫn, ô bảng */
function strip(text: string): string {
  return String(text || "")
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/^\s*>.*$/gm, " ")
    .replace(/^\s*\|.*\|\s*$/gm, " ")
    .replace(/“[^”\n]{0,2000}”/g, " “…” ")
    .replace(/"[^"\n]{0,2000}"/g, ' "…" ')
    .replace(/\[[^\]\n]{0,300}\]/g, " ");
}

/** Tách câu: theo dòng và dấu kết câu (giữ dấu hỏi để nhận ra câu hỏi) */
function sentences(text: string): string[] {
  const out: string[] = [];
  for (const line of strip(text).split(/\n+/)) {
    for (const m of line.split(/(?<=[.!?;])\s+/)) {
      const s = m.replace(/^[\s\-–•*#>✅✔☑️✓\d.)]+/u, "").replace(/\*\*/g, "").trim();
      if (s) out.push(s);
    }
  }
  return out;
}

/** Câu trả lời có khẳng định AI đã thay đổi tài liệu không */
export function findWriteClaim(text: string): ClaimInfo {
  let medium: string | undefined;
  for (const s of sentences(text)) {
    if (/\?\s*$/.test(s)) continue;
    // Phủ định, điều kiện, đề xuất chỉ xét ở phần câu đứng TRƯỚC lời khẳng định ("Đã sửa xong, anh có thể xem lại" vẫn là khẳng định)
    const clear = (m: RegExpExecArray | null) => !!m && !reHedge.test(s.slice(0, m.index + m[0].length));
    if (clear(reFirst.exec(s)) || clear(reEnglish.exec(s))) return { strength: "strong", sentence: s.slice(0, 160) };
    if (!medium && (clear(reInitial.exec(s)) || clear(rePassive.exec(s)))) medium = s.slice(0, 160);
  }
  return medium ? { strength: "medium", sentence: medium } : { strength: "none" };
}

/** Yêu cầu của người dùng có ý nhờ thay đổi tài liệu không (soạn, sửa, thêm, xóa, định dạng…) */
export function hasEditIntent(userText: string): boolean {
  return reIntent.test(String(userText || "").slice(0, 4000));
}

/** Công cụ chỉ đọc mà kết quả là "ai đã sửa gì": sau khi gọi, câu "Điều 5 đã được sửa…" thường là lời kể lại, không phải lời nhận đã làm */
export const DESCRIBING_TOOLS = ["word_list_revisions", "word_compare_versions", "word_list_comments"];

export interface TurnFacts {
  /** Lời người dùng của lượt (chữ họ gõ, không gồm phần add-in ghép thêm) */
  userText: string;
  /** Câu trả lời cuối cùng của AI trong lượt */
  answer: string;
  /** Lượt có ít nhất một công cụ ghi chạy thành công */
  wrote: boolean;
  /** Tên các công cụ đã gọi trong lượt */
  toolsCalled: string[];
  /** Số công cụ ghi đã gọi nhưng lỗi hoặc bị người dùng từ chối */
  failedWrites?: number;
}

export interface ClaimWarning {
  /** Câu hiện trên khung */
  text: string;
  sentence: string;
}

/** Quyết định có cảnh báo không. Chỉ gọi cho lượt ở chế độ tác tử đã kết thúc bình thường (không bị dừng, không lỗi). */
export function unverifiedWriteClaim(f: TurnFacts): ClaimWarning | null {
  if (f.wrote) return null;
  const c = findWriteClaim(f.answer);
  if (c.strength === "none" || !c.sentence) return null;
  const intent = hasEditIntent(f.userText);
  const described = f.toolsCalled.some((t) => DESCRIBING_TOOLS.includes(t));
  // Lời kể lại (đọc sửa đổi của người khác, so phiên bản) mà người dùng không nhờ sửa: không cảnh báo
  if (c.strength === "medium" && (!intent || (described && !/chấp nhận|từ chối|duyệt|xử lý/i.test(f.userText)))) return null;
  if (c.strength === "strong" && described && !intent) return null;
  const why = f.failedWrites
    ? `${f.failedWrites} thao tác ghi của lượt này đều không thành công`
    : f.toolsCalled.length
      ? "lượt này AI chỉ gọi công cụ đọc, không có thao tác ghi nào"
      : "lượt này AI không gọi công cụ nào";
  return {
    sentence: c.sentence,
    text: `Cảnh báo: AI nói đã thay đổi tài liệu (“${c.sentence.length > 90 ? c.sentence.slice(0, 90) + "…" : c.sentence}”) nhưng ${why}. Tài liệu CHƯA được thay đổi.`,
  };
}

/** Tin nhắn gửi khi người dùng bấm "Yêu cầu làm lại bằng công cụ" (là một tin nhắn người dùng mới; lịch sử chỉ thêm) */
export const REDO_WITH_TOOLS =
  "Lượt vừa rồi bạn chưa gọi công cụ ghi nào nên tài liệu chưa thay đổi. Hãy thực hiện đúng yêu cầu trước đó bằng cách GỌI CÔNG CỤ, rồi báo lại chính xác những gì công cụ đã làm. Nếu không gọi được công cụ thì nói rõ là chưa làm.";
