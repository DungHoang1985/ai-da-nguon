/** Kiểm nội dung model miễn phí soạn (thử model miễn phí 23/11/2026, `tasks/viec-le/thu-mien-phi-vbhc/thu-mien-phi-vbhc.md`) — hàm thuần, kiểm thử được:
 *  - từ Latinh không phải tiếng Việt lẫn vào văn bản ("burden", "competencia", "cấmélé", "quórum") và chữ dính nhau ("chủthể");
 *  - thứ trong tuần ghi sai so với ngày ("01/01/2027 (Thứ Ba)" — 01/01/2027 là Thứ Sáu).
 *  Chỉ BÁO để model sửa (công cụ trả `check_now`), không tự sửa: tên riêng, thuật ngữ người dùng đưa vẫn có thể đúng. */

/** Bỏ dấu thanh (huyền, sắc, hỏi, ngã, nặng), giữ dấu mũ, trăng, móc */
const toneless = (w: string) =>
  w
    .normalize("NFD")
    .replace(/[̣̀́̃̉]/g, "")
    .normalize("NFC");

/** Một âm tiết tiếng Việt viết thường: phụ âm đầu? + vần (1–3 nguyên âm) + phụ âm cuối? */
const SYLLABLE = /^(?:ngh|ng|nh|ch|gh|gi|kh|ph|qu|th|tr|[bcdđghklmnpqrstvx])?[aăâeêioôơuưy]{1,3}(?:ch|ng|nh|[cmnpt])?$/u;

/** Từ ngoại thường gặp trong văn bản tiếng Việt, không báo */
const ALLOW = new Set(
  "email website web online offline internet fax wifi file pdf word excel zalo facebook marketing logistics showroom container video audio camera laptop smartphone app server cloud backup password login logout ok vat gps sms usd vnd eur kwh kw km kg cm mm ha".split(
    " "
  )
);

export function isVietSyllable(word: string): boolean {
  return SYLLABLE.test(toneless(word.toLowerCase()));
}

/** Từ viết thường không phải âm tiết tiếng Việt (từ ngoại, hai âm tiết dính nhau). Bỏ qua: từ có chữ hoa (tên riêng, chữ viết tắt),
 *  có chữ số, ngắn hơn 3 chữ, nằm trong ngoặc kép. Tối đa 6 mục dạng `<từ> (sau “<chữ đứng trước>”)`. */
export function strangeWords(lines: string[], max = 6): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const raw of lines) {
    // Bỏ chữ trong ngoặc kép, email, địa chỉ web (soát mã 4.11.11, #9)
    const l = raw
      .normalize("NFC")
      .replace(/[“"][^”"]*[”"]/g, (m) => " ".repeat(m.length))
      .replace(/\S+@\S+|(?:https?:\/\/|www\.)\S+|\b[\w-]+\.(?:com|vn|net|org|gov|edu)(?:\.vn)?\b\S*/giu, (m) => " ".repeat(m.length));
    for (const m of l.matchAll(/[\p{L}\p{M}]+/gu)) {
      const w = m[0];
      if (w.length < 3 || w !== w.toLowerCase() || seen.has(w) || ALLOW.has(w) || isVietSyllable(w)) continue;
      // Chữ Hán, Kirin… đã có foreignRuns báo riêng
      if (!/^[a-zà-ỹđ]+$/u.test(w)) continue;
      seen.add(w);
      const before = Array.from(l.slice(0, m.index)).slice(-20).join("").trim();
      out.push(before ? `${w} (sau “${before}”)` : w);
      if (out.length >= max) return out;
    }
  }
  return out;
}

const THU = ["chủ nhật", "thứ hai", "thứ ba", "thứ tư", "thứ năm", "thứ sáu", "thứ bảy"];
const THU_DISPLAY = ["Chủ nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"];
const THU_RE = "(chủ\\s+nhật|thứ\\s+(?:hai|ba|tư|năm|sáu|bảy|bẩy|2|3|4|5|6|7))";
const DATE_RE = "(?:ngày\\s+)?(\\d{1,2})\\s*(?:/|-|\\.|\\s+tháng\\s+)(\\d{1,2})\\s*(?:/|-|\\.|\\s+năm\\s+)(\\d{4})";

const thuIndex = (t: string): number => {
  const s = t.toLowerCase().replace(/\s+/g, " ").replace("bẩy", "bảy");
  const d = s.match(/thứ (\d)/);
  if (d) return Number(d[1]) - 1;
  return THU.indexOf(s);
};

/** Ngày kèm thứ trong cùng câu: "01/01/2027 (Thứ Ba)", "thứ Hai, ngày 04 tháng 01 năm 2027". Trả câu báo cho từng chỗ sai. */
export function weekdayIssues(lines: string[], max = 6): string[] {
  const out: string[] = [];
  const after = new RegExp(`${DATE_RE}\\s*[(,]?\\s*${THU_RE}`, "giu");
  const before = new RegExp(`${THU_RE}\\s*[,(]?\\s*${DATE_RE}`, "giu");
  for (const raw of lines) {
    const l = raw.normalize("NFC");
    /** "thứ" là số thứ tự, không phải thứ trong tuần (soát mã 4.11.11, #2): "Phiên họp thứ 2, ngày …", "Đại hội lần thứ hai, ngày …";
     *  "thứ" viết bằng số chỉ nhận khi nằm trong ngoặc ngay sau ngày: "15/12/2026 (thứ 3)" */
    const ordinal = (thu: string, idx: number, inParen: boolean) => {
      if (/\d/.test(thu) && !inParen) return true;
      if (/^chủ/iu.test(thu)) return false;
      const prev = l.slice(Math.max(0, idx - 24), idx).trim().toLocaleLowerCase("vi-VN");
      return /(?:^|[^\p{L}])(?:lần|kỳ|kì|phiên|khóa|khoá|đợt|lượt|vòng|bước|họp|hội|hội nghị|điều|khoản|số|hạng|tầng|người|lớp|năm)$/u.test(prev);
    };
    const check = (dd: string, mm: string, yy: string, thu: string, at: string) => {
      const d = Number(dd);
      const m = Number(mm);
      const y = Number(yy);
      const dt = new Date(Date.UTC(y, m - 1, d));
      if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return;
      const want = dt.getUTCDay();
      const got = thuIndex(thu);
      if (got < 0 || got === want) return;
      const msg = `${String(d).padStart(2, "0")}/${String(m).padStart(2, "0")}/${y} là ${THU_DISPLAY[want]}, văn bản ghi “${at.trim()}”`;
      if (!out.includes(msg)) out.push(msg);
    };
    for (const m of l.matchAll(after)) {
      const k = m[0].length - m[4].length;
      if (!ordinal(m[4], (m.index ?? 0) + k, /\(\s*$/.test(m[0].slice(0, k)))) check(m[1], m[2], m[3], m[4], m[0]);
    }
    for (const m of l.matchAll(before)) if (!ordinal(m[1], m.index ?? 0, false)) check(m[2], m[3], m[4], m[1], m[0]);
    if (out.length >= max) return out.slice(0, max);
  }
  return out;
}

const dayNo = (dd: string, mm: string, yy: string): number | null => {
  const d = Number(dd);
  const m = Number(mm);
  const dt = new Date(Date.UTC(Number(yy), m - 1, d));
  return dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d ? dt.getTime() : null;
};
const fmtDay = (t: number) => {
  const d = new Date(t);
  return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
};

/** Giấy mời ghi ngày ban hành trùng hoặc sau ngày họp (thử 24/11/2026, dots-3-note: lấy ngày họp làm ngày ban hành).
 *  Ngày họp = ngày trong dòng có giờ ("14h00", "8 giờ 30") hoặc dòng "Thời gian". Chỉ báo khi đọc được cả hai ngày. */
export function issueAfterEvent(issueDate: string, lines: string[]): string | null {
  const re = new RegExp(DATE_RE, "iu");
  const im = re.exec(issueDate.normalize("NFC"));
  const issue = im ? dayNo(im[1], im[2], im[3]) : null;
  if (issue === null) return null;
  for (const raw of lines) {
    const l = raw.normalize("NFC");
    // Giờ họp: "14h00", "8 giờ 30" — \b của JS chỉ hiểu chữ ASCII nên chặn bằng (?!\p{L}); bỏ dòng căn cứ (soát mã 4.11.11, #6)
    if (/^[\s_*#-]*(?:căn cứ|thực hiện|theo|xét)(?!\p{L})/iu.test(l)) continue;
    if (!/\d{1,2}\s*(?:h|giờ)(?:\s*\d{1,2})?(?![\p{L}\d])|thời gian/iu.test(l)) continue;
    const m = re.exec(l);
    const ev = m ? dayNo(m[1], m[2], m[3]) : null;
    if (ev !== null && ev <= issue)
      return `Ngày ban hành ${fmtDay(issue)} ${ev === issue ? "trùng" : "sau"} ngày họp ${fmtDay(ev)}: giấy mời phải ban hành TRƯỚC ngày họp — sửa ngày ban hành hoặc để trống cho người dùng điền`;
  }
  return null;
}

/** Văn bản đã hết hiệu lực model hay dẫn theo trí nhớ (thử 23–24/11/2026: Luật Nhà ở 2014, Luật Đất đai 2013, Nghị định 99/2015 —
 *  kể cả khi kỹ năng đã dặn). Mã kiểm theo tên + năm hoặc số hiệu; chỉ báo, không sửa. Soát lại danh mục khi luật đổi. */
const EXPIRED: [RegExp, string][] = [
  [/Luật Nhà ở(?: số 65\/2014\/QH13| năm 2014| 2014)|65\/2014\/QH13/iu, "Luật Nhà ở 2014 (thay bằng Luật Nhà ở năm 2023)"],
  [/Luật Đất đai(?: số 45\/2013\/QH13| năm 2013| 2013)|45\/2013\/QH13/iu, "Luật Đất đai 2013 (thay bằng Luật Đất đai năm 2024)"],
  [/Luật Kinh doanh bất động sản(?: số 66\/2014\/QH13| năm 2014| 2014)|66\/2014\/QH13/iu, "Luật Kinh doanh bất động sản 2014 (thay bằng luật năm 2023)"],
  [/Luật Doanh nghiệp(?: số 68\/2014\/QH13| năm 2014| 2014)|68\/2014\/QH13/iu, "Luật Doanh nghiệp 2014 (thay bằng Luật Doanh nghiệp năm 2020)"],
  [/99\/2015\/NĐ-CP|Nghị định(?: số)? 99\/2015/iu, "Nghị định 99/2015/NĐ-CP hướng dẫn Luật Nhà ở 2014"],
  [/02\/2022\/NĐ-CP|Nghị định(?: số)? 02\/2022/iu, "Nghị định 02/2022/NĐ-CP hướng dẫn Luật Kinh doanh bất động sản 2014"],
];
export function expiredLaws(lines: string[]): string[] {
  const text = lines.join("\n").normalize("NFC");
  return EXPIRED.filter(([re]) => re.test(text)).map(([, name]) => name);
}

/** Chữ model vừa ghi khi SỬA văn bản có lẫn từ lạ mà chữ cũ không có (vòng thử 3, 24/11/2026: dots-3-note ghi "DeniedĐến nay",
 *  "Collaboratively配合", "Unauthorized ban nhân dân"; trước đây chỉ kiểm lúc dựng văn bản mới). Trả câu `check_now` hoặc null.
 *  Từ lạ dính liền chữ hoa ("DeniedĐến") cũng bắt: tách tại chỗ chữ thường → chữ hoa. */
export function newTextCheck(newText: string, oldText = ""): string | null {
  const split = (s: string) => s.normalize("NFC").replace(/(\p{Ll})(\p{Lu})/gu, "$1 $2");
  const had = new Set(strangeWords([split(oldText)], 50).map((x) => x.split(" ")[0]));
  // Từ Latinh viết hoa chỉ tính khi DÍNH LIỀN chữ hoa tiếng Việt phía sau ("DeniedĐến"); đứng riêng có thể là tên riêng ("Riverside")
  const latin = (s: string) =>
    (s.normalize("NFC").match(/(?<![\p{L}])[A-Z][a-z]{3,}(?=[ĐÂĂÊÔƠƯÁÀẢÃẠẤẦẨẪẬẮẰẲẴẶÉÈẺẼẸẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌỐỒỔỖỘỚỜỞỠỢÚÙỦŨỤỨỪỬỮỰÝỲỶỸỴ])/gu) || []).filter((w) => !isVietSyllable(w.toLowerCase()));
  const oldLatin = new Set(latin(oldText).map((w) => w.toLowerCase()));
  const odd = [...new Set([...strangeWords([split(newText)], 6).map((x) => x.split(" ")[0]), ...latin(newText)])].filter((w) => !had.has(w) && !oldLatin.has(w.toLowerCase()));
  const han = (newText.match(/[\p{Script=Han}\p{Script=Hangul}\p{Script=Cyrillic}\p{Script=Arabic}]+/gu) || []).filter((w) => !oldText.includes(w));
  const all = [...odd, ...han];
  if (!all.length) return null;
  return `Chữ vừa ghi có từ không phải tiếng Việt: ${all.slice(0, 6).join(", ")} — trừ khi là tên riêng, sửa NGAY bằng word_replace_text (find = đúng cụm đó) trước khi trả lời`;
}

/** Thẻ trích dẫn Cohere lọt vào tham số ("<co>Thông qua</co: 2:[0]>"), ký hiệu markdown đậm; trả chuỗi đã bỏ */
export function stripModelMarkup(s: string): string {
  return (
    s
      .replace(/<\/?co(?::[^>]*)?>/g, "")
      // Dấu nhấn markdown dính liền chữ thay cho khoảng trắng: "Quân*uỷ*quyền", "),*uỷ*quyền" (vòng thử 3, 24/11/2026: dots-3-note).
      // Chỉ MỘT từ (không khoảng trắng) kẹp giữa hai dấu * và dính chữ ở sau — chữ đậm hợp lệ "**Ông A**, Chức vụ" không bị đụng (soát mã 4.11.11, #7)
      .replace(/(?<=[\p{L},;:)])\*{1,2}(\p{L}+)\*{1,2}(?=\p{L})/gu, " $1 ")
      .replace(/ {2,}/g, " ")
  );
}
