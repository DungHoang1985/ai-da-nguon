/** Nhận diện ca "AI nói đã sửa nhưng lượt đó không có thao tác ghi nào" (Đợt 3, việc A2) — hàm thuần, kiểm thử được.
 *  Kết quả chỉ dùng để HIỆN cảnh báo trên khung; không đi vào lịch sử gửi model, không đổi lời nhắc hệ thống.
 *  Bối cảnh: 16/11/2026 một model miễn phí trả lời "Đã sửa…" sau 3 giây mà không gọi công cụ nào, văn bản không đổi.
 *  Ưu tiên KHÔNG báo nhầm (cảnh báo sai nhiều thì người dùng sẽ quen mắt và bỏ qua): chỉ bắt lời khẳng định rõ là đã thay đổi tài liệu.
 *  Các nhóm báo nhầm đã gặp khi soát mã độc lập (văn cảnh pháp lý) có kiểm thử chặn trong test/unit/dot3.test.ts. */

const L = "[^\\p{L}\\p{N}]";
const END = `(?=$|${L})`;

/** Động từ chỉ việc thay đổi tài liệu. Từ đa nghĩa có điều kiện đi kèm: "ghi" (không phải ghi nhận, ghi chú), "chuyển" (không phải
 *  chuyển nhượng, chuyển sang đọc…), "căn" (chỉ căn lề / giữa…), "đặt" (chỉ đặt lề, đặt lại…), "bỏ" (không phải bỏ qua), "thay" (không phải thay mặt). */
const VERBS = [
  "sửa(?!\\s+đổi,?\\s+bổ\\s+sung)",
  "chỉnh sửa",
  "chỉnh(?!\\s+phủ)",
  "thay thế",
  "thay(?!\\s+mặt)",
  "thêm",
  "chèn",
  "xóa",
  "xoá",
  "ghi(?!\\s+(?:nhận|chú|rõ\\s+ở|lại\\s+ý))",
  "cập nhật",
  "đổi",
  "điền",
  "nhập",
  "soạn",
  "tạo",
  "dựng",
  "định dạng",
  "in đậm",
  "in nghiêng",
  "gạch chân",
  "viết lại",
  "viết",
  "bổ sung",
  "áp dụng(?=\\s+(?:định dạng|kiểu|mẫu|bố cục|thay đổi|sửa đổi))",
  "di chuyển",
  "chuyển(?=\\s+(?:đoạn|điều|khoản|slide|hàng|cột|ô|bảng|hình|nội dung|chữ|văn bản|mục)|\\s+\\S+\\s+(?:lên|xuống|ra sau|ra trước))",
  "đặt(?=\\s+(?:lại|lề|khổ|cỡ|phông|màu|tên|tiêu đề|bố cục|định dạng))",
  "căn(?=\\s+(?:lề|giữa|trái|phải|đều|chỉnh))",
  "tô",
  "gỡ",
  "bỏ(?!\\s+qua)",
  "loại bỏ",
  "đánh dấu",
  "chấp nhận",
  "từ chối",
  "trả lời nhận xét",
  "sắp xếp",
  "gộp",
  "tách",
  "nhân bản",
  "điều chỉnh",
  "làm lại",
  "trình bày lại",
  "vẽ",
].join("|");
const VERB = `(?:${VERBS})`;
/** Chủ ngữ là chính AI ("chúng tôi", "ai" không tính) */
const FIRST = "(?<!chúng\\s)(?:tôi|em|mình|tớ)";
/** Chủ ngữ là một phần của tài liệu (câu bị động: "Đoạn 3 đã được sửa") */
const OBJECT =
  "(?:văn bản|tài liệu|hợp đồng|công văn|tờ trình|đoạn|điều|khoản|điểm|mục|câu|ô|bảng|sheet|trang tính|sổ tính|cột|hàng|dòng|công thức|biểu đồ|slide|trang|bài trình chiếu|tiêu đề|nội dung|số liệu|dữ liệu|tệp|file|nhận xét|định dạng|lề|phông|chữ|mức phạt|giá trị)";
const MID = "(?:[^.;:!?\\n]|\\.(?=\\d))";

const reFirst = new RegExp(`(?:^|${L})${FIRST}\\s+(?:cũng\\s+|vừa\\s+)?(?:đã|vừa)\\s+(?:tiến hành\\s+|thực hiện\\s+(?:việc\\s+)?|hoàn (?:tất|thành)\\s+việc\\s+)?${VERB}${END}`, "iu");
/** "Tôi đã … và / rồi / nhưng đã sửa …": chủ ngữ ở vế trước */
const reFirstLater = new RegExp(`(?:^|${L})${FIRST}\\s+đã${L}${MID}{0,80}?(?:${L})(?:và|rồi|nhưng|đồng thời)\\s+(?:cũng\\s+)?đã\\s+${VERB}${END}`, "iu");
const reInitial = new RegExp(`^(?:đã|vừa)\\s+(?:tiến hành\\s+|thực hiện\\s+(?:việc\\s+)?|hoàn (?:tất|thành)\\s+việc\\s+)?${VERB}${END}`, "iu");
const rePassive = new RegExp(`(?:^|${L})${OBJECT}${MID}{0,60}?\\s(?:đã|vừa)\\s+(?:được|bị)\\s+(?:tôi\\s+)?${VERB}${END}`, "iu");
/** "Điều 5: đã sửa" */
const reColon = new RegExp(`(?:^|${L})${OBJECT}${MID}{0,40}?:\\s*(?:đã|vừa)\\s+${VERB}${END}`, "iu");
const reEnglish = /\b(?:i(?:'ve| have)|i)\s+(?:now\s+|just\s+|already\s+)?(?:updated|changed|replaced|inserted|added|deleted|removed|edited|fixed|modified|applied|rewritten|rewrote|formatted|corrected)\b|\b(?:has|have) been (?:updated|changed|replaced|inserted|added|deleted|removed|edited|fixed|modified|applied|rewritten|formatted|corrected)\b/i;

/** Vế câu chứa lời khẳng định mà có các từ này đứng trước thì không phải khẳng định: phủ định, tương lai, điều kiện, lời kể về người khác */
const reHedge = new RegExp(
  `(?:^|${L})(?:chưa|không thể|không\\s+\\p{L}+\\s+được|đã không|chẳng|sẽ|nếu|giả sử|khi nào|một khi|có thể|hãy|bạn đã|anh đã|chị đã|người dùng đã|đối tác|bên\\s+[a-zđ]+\\s+đã|khách hàng|tác giả|ai đã|ai đó)${END}`,
  "iu"
);
/** Câu là lời kể lại lịch sử của văn bản hoặc dẫn vào phần trình bày trong khung chat, không phải nhận đã làm */
const reNarrative = new RegExp(
  `(?:^|${L})(?:dưới đây|bên dưới|sau đây|như sau|bởi|nhiều lần|trước đây|trước đó|từ trước|lần trước|ở lượt trước|theo\\s+(?:phụ lục|thỏa thuận|thoả thuận|luật|quy định|nghị định|thông tư|văn bản|bản)|năm\\s+\\d{4}|ngày\\s+\\d{1,2}[/-]\\d)${END}`,
  "iu"
);

/** Ý nhờ thay đổi tài liệu trong lời người dùng: chỉ các động từ rõ nghĩa (không tính "chuyển nhượng", "căn hộ", "đặt cọc", "ghi chú"…) */
const reIntent = new RegExp(
  `(?:^|${L})(?:sửa|chỉnh|thay(?!\\s+mặt)|thêm|chèn|xóa|xoá|cập nhật|đổi|điền|soạn|dựng|định dạng|in đậm|viết|bổ sung|tô|gỡ|chấp nhận|từ chối|trả lời nhận xét|sắp xếp|gộp|tách|rút gọn|dịch|chuẩn hóa|chuẩn hoá|xử lý|hoàn thiện|làm lại|tạo|lập|vẽ|ghi(?!\\s+(?:nhận|chú))|chuyển(?!\\s+(?:nhượng|giao|tiền|khoản|đổi|quyền))|đặt(?=\\s+(?:lại|lề|khổ|cỡ|phông|màu|tên))|căn(?=\\s+(?:lề|giữa|trái|phải|đều|chỉnh))|GỌI CÔNG CỤ|fix|edit|update|rewrite|replace|insert|add|delete|remove|format|write|create|change)${END}`,
  "iu"
);
/** Câu hỏi "có cần sửa không", "đã sửa chưa" không phải lời nhờ sửa */
const reAskOnly = new RegExp(`(?:có\\s+(?:cần|nên|phải)\\s+\\p{L}+|đã\\s+(?:được\\s+)?\\p{L}+(?:\\s+\\p{L}+)?\\s+chưa)[^.!?]*\\?\\s*$`, "iu");

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

/** Tách câu: theo dòng và dấu kết câu đứng trước khoảng trắng (giữ "5.3", "0,05%") */
function sentences(text: string): string[] {
  const out: string[] = [];
  for (const line of strip(text).split(/\n+/)) {
    for (const m of line.split(/(?<=[.!?;])\s+/)) {
      const s = m
        .replace(/^[\s\-–•*#>✅✔☑️✓\d.)]+/u, "")
        .replace(/\*\*/g, "")
        .trim();
      if (s) out.push(s);
    }
  }
  return out;
}

/** Vế câu chứa vị trí `at`: tính từ dấu phẩy / hai chấm / gạch ngang gần nhất phía trước */
function clauseBefore(s: string, at: number): string {
  const head = s.slice(0, at);
  const cut = Math.max(head.lastIndexOf(","), head.lastIndexOf(":"), head.lastIndexOf(" – "), head.lastIndexOf(" — "), head.lastIndexOf(" vì "), head.lastIndexOf(" nhưng "));
  return head.slice(cut + 1);
}

/** Câu trả lời có khẳng định AI đã thay đổi tài liệu không */
export function findWriteClaim(text: string): ClaimInfo {
  let medium: string | undefined;
  for (const s of sentences(text)) {
    if (/\?\s*$/.test(s) || reNarrative.test(s)) continue;
    // Phủ định, điều kiện, lời kể về người khác chỉ xét trong vế câu đứng ngay trước lời khẳng định
    const clear = (m: RegExpExecArray | null) => !!m && !reHedge.test(clauseBefore(s, m.index) + m[0]);
    // "Tôi đã không thể sửa Điều 5 nhưng đã sửa Điều 6": chỉ xét vế sau từ nối
    const later = reFirstLater.exec(s);
    const laterOk = !!later && !reHedge.test(later[0].split(/\s(?:và|rồi|nhưng|đồng thời)\s/).pop() || "");
    if (clear(reFirst.exec(s)) || laterOk || clear(reEnglish.exec(s))) return { strength: "strong", sentence: s.slice(0, 160) };
    if (!medium && (clear(reInitial.exec(s)) || clear(rePassive.exec(s)) || clear(reColon.exec(s)))) medium = s.slice(0, 160);
  }
  return medium ? { strength: "medium", sentence: medium } : { strength: "none" };
}

/** Yêu cầu của người dùng có ý nhờ thay đổi tài liệu không (soạn, sửa, thêm, xóa, định dạng…) */
export function hasEditIntent(userText: string): boolean {
  const t = String(userText || "").slice(0, 4000);
  return reIntent.test(t) && !reAskOnly.test(t.trim());
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
  /** Số công cụ ghi đã gọi nhưng lỗi, không đổi được gì, hoặc bị người dùng từ chối */
  failedWrites?: number;
  /** Lượt này nối tiếp nhiệm vụ của lượt trước (người dùng bấm "Làm tiếp"); các trường trên đã gộp cả nhiệm vụ (`mergeTaskFacts`) */
  continued?: boolean;
}

/** Những gì đã làm của một nhiệm vụ tới hết một lượt — lưu ở mục trả lời để lượt "Làm tiếp" kế đó gộp vào */
export interface TaskSoFar {
  /** Yêu cầu gốc người dùng gõ (lượt Làm tiếp chỉ gửi câu của add-in) */
  userText: string;
  wrote: boolean;
  toolsCalled: string[];
  failedWrites: number;
}

/** Gộp lượt hiện tại với các lượt trước của CÙNG nhiệm vụ (22/11/2026: lượt sau "Làm tiếp" chỉ đọc lại rồi báo "đã thêm…" bị cảnh báo nhầm,
 *  dù thao tác ghi đã chạy ở lượt trước — X6-D3, W4-B3). Không có `prev` (nhiệm vụ mới) thì trả nguyên lượt hiện tại. */
export function mergeTaskFacts(prev: TaskSoFar | undefined, cur: TurnFacts): TurnFacts {
  if (!prev) return cur;
  return {
    userText: prev.userText || cur.userText,
    answer: cur.answer,
    wrote: prev.wrote || cur.wrote,
    toolsCalled: [...prev.toolsCalled, ...cur.toolsCalled],
    failedWrites: (prev.failedWrites || 0) + (cur.failedWrites || 0),
    continued: true,
  };
}

/** Bản lưu của nhiệm vụ sau lượt này (để lượt Làm tiếp kế đó gộp) */
export function taskSoFar(f: TurnFacts): TaskSoFar {
  return { userText: f.userText, wrote: f.wrote, toolsCalled: f.toolsCalled.slice(-200), failedWrites: f.failedWrites || 0 };
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
  // Câu không có chủ ngữ "tôi" chỉ tính khi người dùng có nhờ sửa; đã đọc sửa đổi của người khác thì phải là nhờ duyệt / xử lý
  if (c.strength === "medium" && (!intent || (described && !/chấp nhận|từ chối|duyệt|xử lý/i.test(f.userText)))) return null;
  if (c.strength === "strong" && described && !intent) return null;
  const scope = f.continued ? "cả nhiệm vụ (tính cả các lượt trước khi bấm Làm tiếp)" : "lượt này";
  const why = f.failedWrites
    ? `${f.failedWrites} thao tác ghi của ${scope} đều không thành công`
    : f.toolsCalled.length
      ? `${f.continued ? "trong " : ""}${scope} AI chỉ gọi công cụ đọc, không có thao tác ghi nào`
      : `${f.continued ? "trong " : ""}${scope} AI không gọi công cụ nào`;
  return {
    sentence: c.sentence,
    text: `Cảnh báo: AI nói đã thay đổi tài liệu (“${c.sentence.length > 90 ? c.sentence.slice(0, 90) + "…" : c.sentence}”) nhưng ${why}. Tài liệu CHƯA được thay đổi.`,
  };
}

/** AI đổ lỗi cho công cụ khi lượt đó không gọi công cụ nào (thử 23/11/2026: ERNIE hai lần nói "công cụ soạn thảo đang lỗi" rồi viết
 *  hợp đồng ra khung chat, không hề gọi công cụ). Chỉ bắt câu khẳng định công cụ hỏng; không bắt khi có công cụ lỗi thật. */
const reToolBroken = new RegExp(
  `(?:^|${L})(?:công cụ|chức năng|tính năng)(?:\\s+[\\p{L}]+){0,3}\\s+(?:(?:hiện\\s+)?(?:đang|vẫn|bị|vẫn đang)\\s+)?(?:lỗi|gặp lỗi|gặp sự cố|không hoạt động|không khả dụng|không phản hồi|tạm ngưng)${END}|(?:^|${L})không thể (?:tạo|dựng|ghi|soạn)\\s+văn bản\\s+(?:ngay\\s+)?(?:lúc này|bây giờ)`,
  "iu"
);
export function toolFailureClaim(f: TurnFacts & { toolErrors?: number }): ClaimWarning | null {
  if (f.toolsCalled.length || f.failedWrites || f.toolErrors) return null;
  const s = sentences(f.answer).find((x) => reToolBroken.test(x) && !/\?\s*$/.test(x));
  if (!s) return null;
  return {
    sentence: s,
    text: `Cảnh báo: AI nói công cụ bị lỗi (“${s.length > 90 ? s.slice(0, 90) + "…" : s}”) nhưng ${f.continued ? "trong cả nhiệm vụ" : "lượt này"} AI không gọi công cụ nào — công cụ không hề lỗi. Bấm "Yêu cầu làm lại bằng công cụ" hoặc đổi model.`,
  };
}

/** Công cụ ghi có thật sự thay đổi tài liệu không: lỗi, trả ok=false, hoặc báo thay 0 chỗ đều không tính */
export function writeSucceeded(result: any, error?: string): boolean {
  if (error) return false;
  if (result && typeof result === "object") {
    if (result.ok === false) return false;
    if (result.replaced === 0) return false;
  }
  return true;
}

/** Tin nhắn gửi khi người dùng bấm "Yêu cầu làm lại bằng công cụ" (là một tin nhắn người dùng mới; lịch sử chỉ thêm) */
export const REDO_WITH_TOOLS =
  "Lượt vừa rồi bạn chưa gọi công cụ ghi nào nên tài liệu chưa thay đổi. Hãy thực hiện đúng yêu cầu trước đó bằng cách GỌI CÔNG CỤ, rồi báo lại chính xác những gì công cụ đã làm. Nếu không gọi được công cụ thì nói rõ là chưa làm.";
