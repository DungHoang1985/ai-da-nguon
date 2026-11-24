/** Dựng bố cục văn bản hành chính (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản…) đúng thể thức:
 *  hàm thuần, biến các trường AI đưa thành một "bản vẽ" gồm bảng không viền và các đoạn đã có sẵn cỡ chữ, kiểu chữ, căn lề.
 *  Phần Office.js (wordAdmin.ts) chỉ việc dựng theo bản vẽ, nên thể thức đúng dù model mạnh hay yếu.
 *  Số đo theo chuẩn trình bày văn bản hành chính đang áp dụng: A4, Times New Roman, cỡ chữ 13–14. */
import { expiredLaws, strangeWords, stripModelMarkup, weekdayIssues } from "./textcheck";

export const DOC_TYPES = ["cong_van", "to_trinh", "bao_cao", "quyet_dinh", "nghi_quyet", "thong_bao", "giay_moi", "bien_ban", "ke_hoach", "hop_dong", "khac"] as const;
export type DocType = (typeof DOC_TYPES)[number];

/** Tên loại in trên văn bản (công văn không có tên loại) */
export const TYPE_TITLE: Record<DocType, string> = {
  cong_van: "",
  to_trinh: "TỜ TRÌNH",
  bao_cao: "BÁO CÁO",
  quyet_dinh: "QUYẾT ĐỊNH",
  nghi_quyet: "NGHỊ QUYẾT",
  thong_bao: "THÔNG BÁO",
  giay_moi: "GIẤY MỜI",
  bien_ban: "BIÊN BẢN",
  ke_hoach: "KẾ HOẠCH",
  hop_dong: "HỢP ĐỒNG",
  khac: "",
};

/** Lề chuẩn (mm) và khổ giấy khi công cụ đặt thể thức trang */
export const ADMIN_PAGE = { paper: "A4", topMm: 20, bottomMm: 20, leftMm: 30, rightMm: 20 } as const;
/** Bề rộng vùng trình bày với lề chuẩn (mm) */
export const TEXT_WIDTH_MM = 210 - ADMIN_PAGE.leftMm - ADMIN_PAGE.rightMm;

export type Align = "Left" | "Centered" | "Justified" | "Right";

export interface Para {
  text: string;
  size: number;
  bold?: boolean;
  italic?: boolean;
  align: Align;
  /** Thụt đầu dòng 10 mm (đoạn nội dung) */
  indent?: boolean;
  /** Cách đoạn phía sau (pt) */
  after?: number;
  /** Cách đoạn phía trước (pt) */
  before?: number;
  /** Giãn dòng theo số dòng (1 = đơn) */
  line?: number;
  /** Thụt cả đoạn về bên trái (mm) */
  leftMm?: number;
  /** Đường kẻ ngang thay cho chữ: bề dài (mm), đặt giữa */
  ruleMm?: number;
  keepNext?: boolean;
  /** Chỉ phần đầu này in đậm, phần sau chữ thường ("Điều 1." của quyết định, nghị quyết) */
  boldLead?: string;
}

export interface TableBlock {
  kind: "table";
  /** Bề rộng hai cột (mm) */
  widths: [number, number];
  /** Mỗi hàng hai ô, mỗi ô là một dãy đoạn (ít nhất một). Đã đo trong Word 16/11/2026: thêm đoạn vào ô phải chèn "Before" đoạn sẵn có
   *  ("After" rơi sang ô kế); viền dưới của đoạn DUY NHẤT hoặc đoạn CUỐI trong ô dính sang ô cùng hàng → đường kẻ không bao giờ là đoạn cuối ô. */
  rows: [Para[], Para[]][];
}
export interface ParaBlock extends Para {
  kind: "para";
}
export type Block = TableBlock | ParaBlock;

export interface AdminDocInput {
  doc_type?: string;
  /** Tên cơ quan, tổ chức chủ quản trực tiếp (nếu có) */
  parent_org?: string;
  /** Tên cơ quan, tổ chức ban hành */
  org?: string;
  /** Số văn bản; bỏ trống để văn thư cấp */
  number?: string;
  /** Ký hiệu, ví dụ "TTr-ABC", "ABC-PC" */
  symbol?: string;
  /** Địa danh */
  place?: string;
  /** Ngày tháng năm, ví dụ "ngày 05 tháng 01 năm 2026"; bỏ trống để chừa chỗ */
  date?: string;
  /** Tên loại khi doc_type = khac hoặc muốn ghi khác mặc định */
  title?: string;
  /** Trích yếu (với công văn là nội dung sau "V/v") */
  subject?: string;
  /** Kính gửi */
  to?: string[] | string;
  /** Nội dung: mỗi dòng một đoạn. "# " = đề mục in đậm canh giữa; "## " = đề mục in đậm; "_…_" cả dòng = chữ nghiêng (căn cứ) */
  body?: string;
  /** Quyền hạn ký, ví dụ "TM. HỘI ĐỒNG QUẢN TRỊ", "KT. GIÁM ĐỐC" */
  signer_authority?: string;
  /** Chức vụ người ký */
  signer_title?: string;
  signer_name?: string;
  /** Người ký bên trái (biên bản: thư ký) */
  left_signer_title?: string;
  left_signer_name?: string;
  /** Hợp đồng: dòng đầu khối ký bên trái (mặc định "ĐẠI DIỆN BÊN A") */
  left_signer_authority?: string;
  /** Nơi nhận */
  recipients?: string[] | string;
  /** Cỡ chữ nội dung: 13 hoặc 14 */
  font_size?: number;
  /** Năm hiện tại (để chừa chỗ ngày tháng) */
  year?: number;
}

export interface AdminDocPlan {
  type: DocType;
  blocks: Block[];
  /** Những chỗ còn để trống, cần điền trước khi ký */
  blanks: string[];
  /** Lưu ý về nội dung do công cụ tự chỉnh */
  notes: string[];
  size: number;
  /** Cụm chữ lạ (Hán, Nhật, Ả Rập, Kirin) trong nội dung — model miễn phí đôi khi chèn lẫn (Chạy thật 23/11/2026) */
  foreign?: string[];
  /** Chỗ nội dung nghi sai để model kiểm và sửa ngay: thứ trong tuần lệch ngày, từ không phải tiếng Việt, quyền hạn ký sai dạng… */
  checks?: string[];
}

const clean = (s: unknown) =>
  stripModelMarkup(String(s ?? ""))
    .replace(/\s+/g, " ")
    .trim();
/** Danh sách nơi nhận: mảng giữ nguyên từng phần tử; chuỗi thì tách theo dòng hoặc dấu chấm phẩy. Bỏ gạch đầu dòng và dấu câu cuối,
 *  nhưng giữ dấu chấm của chữ viết tắt ("UBND TP.", "Co., Ltd."). */
const list = (v: unknown): string[] =>
  (Array.isArray(v) ? v.filter((x) => typeof x === "string" || typeof x === "number") : String(v ?? "").split(/\r?\n|;\s*(?=\S)/))
    .map((x) =>
      clean(x)
        .replace(/^[-–•+]\s*/, "")
        .replace(/[;,]+$/, "")
        .replace(/(\p{L}{4,}|\d)\.+$/u, "$1")
        .replace(/[;,]+$/, "")
    )
    .filter(Boolean);
const noAccent = (s: string) =>
  s
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");
const upper = (s: string) => s.toLocaleUpperCase("vi-VN");
const BLANK = "…";
/** Dấu chấm kết thúc danh sách; mục đã kết bằng dấu chấm của chữ viết tắt thì không thêm */
const dot = (t: string) => (t.endsWith(".") ? "" : ".");

/** Bề rộng ước lượng của một dòng chữ Times New Roman (mm) — để tính độ dài đường kẻ */
export function estimateWidthMm(text: string, size: number, o: { bold?: boolean } = {}): number {
  let em = 0;
  for (const ch of text) {
    if (ch === " ") em += 0.25;
    else if (/[.,;:'’!\-–()/]/.test(ch)) em += 0.3;
    else if (/[ilIítjfr]/.test(ch)) em += 0.31;
    else if (/[mwMW]/.test(ch)) em += 0.86;
    else if (ch !== ch.toLocaleLowerCase("vi-VN")) em += 0.7;
    else em += 0.47;
  }
  return (em * size * (o.bold ? 1.04 : 1) * 25.4) / 72;
}

const QUOC_HIEU = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
const TIEU_NGU = "Độc lập - Tự do - Hạnh phúc";

/** Nội dung (body) thành các dòng. Model miễn phí đưa đủ kiểu (thử 23/11/2026): mảng; chuỗi là MẢNG JSON
 *  ('["_Căn cứ …_", "## Điều 1. …"]' — Nemotron 3 Ultra, cả văn bản thành một đoạn); chuỗi chỉ có "\n" viết bằng chữ. */
export function bodyLinesOf(body: unknown): string[] {
  let v: unknown = body;
  if (typeof v === "string") {
    const t = v.trim();
    if (/^\[\s*"/.test(t) && /"\s*,?\s*\]$/.test(t)) {
      try {
        const arr = JSON.parse(t);
        if (Array.isArray(arr)) v = arr;
      } catch {
        // Mảng hỏng (thiếu dấu ngoặc kép…): tách tay theo '", "'
        v = t
          .replace(/^\[\s*"/, "")
          .replace(/"\s*,?\s*\]$/, "")
          .split(/"\s*,\s*"/);
      }
    }
  }
  const text = Array.isArray(v) ? (v as unknown[]).map((x) => String(x ?? "")).join("\n") : String(v ?? "");
  // "\n" viết bằng chữ, kể cả lẫn giữa các dòng thật (thử 24/11/2026: Nemotron Nano "THẨM QUYỀN BAN HÀNH\n## Điều 1. …")
  const lines = stripModelMarkup(text).replace(/\\n/g, "\n").split(/\r?\n/).map((l) => l.trim());
  // Một cặp _…_ bao nhiều dòng ("_Căn cứ Điều lệ;" … "… tán thành._" — dots-3-note): nghiêng từng dòng, không in dấu gạch dưới
  for (let i = 0; i < lines.length; i++) {
    if (!/^_[^_\s]/.test(lines[i]) || /_$/.test(lines[i])) continue;
    const j = lines.findIndex((l, k) => k > i && /[^_]_$/.test(l) && !/^_/.test(l));
    if (j < 0 || j - i > 12) continue;
    for (let k = i; k <= j; k++) if (lines[k]) lines[k] = `_${lines[k].replace(/^_/, "").replace(/_$/, "")}_`;
    i = j;
  }
  // Mở "_" mà không đóng ("_Căn cứ Luật Nhà ở năm 2023;" — Nemotron 3 Ultra, kiểm chứng 24/11/2026): đóng lại cho thành dòng nghiêng
  for (let i = 0; i < lines.length; i++) if (/^_[^_\s]/.test(lines[i]) && !/_$/.test(lines[i])) lines[i] += "_";
  return lines
    .join("\n")
    .replace(/\\n/g, "\n")
    .split(/\r?\n/)
    .map((l) => l.trim())
    // "Căn cứ Điều lệ công ty.;" → "…công ty;"
    .map((l) => l.replace(/\.\s*;(_?)$/, ";$1"))
    // Gạch đầu dòng bằng dấu gạch dưới ("_ Tên khoản vay: …" — Cohere) thành gạch ngang; dòng nghiêng _…_ giữ nguyên
    .map((l) => (/^_\s+/.test(l) && !/_$/.test(l) ? "- " + l.replace(/^_\s+/, "") : l))
    .filter(Boolean);
}

/** Tách một dòng nội dung thành đoạn có định dạng */
function bodyPara(raw: string, size: number): Para {
  const line = raw.trim();
  const h1 = line.match(/^#\s+(.*)$/);
  if (h1) return { text: h1[1].replace(/\*\*/g, ""), size, bold: true, align: "Centered", after: 6, before: 6, line: 1.2, keepNext: true };
  const h2 = line.match(/^#{2,4}\s+(.*)$/) || line.match(/^\*\*(.+)\*\*$/);
  if (h2) {
    const text = h2[1].replace(/\*\*/g, "");
    // "Điều 1. Bổ nhiệm ông … kể từ ngày …": chỉ "Điều 1." in đậm, phần lời văn chữ thường (thử 23/11/2026: cả điều in đậm).
    // Điều có tên ngắn ("Điều 1. Phạm vi điều chỉnh") vẫn đậm cả dòng như tên điều.
    const lead = text.match(/^(Điều\s+\d+[a-zđ]?\s*[.:])\s+(.+)$/iu);
    if (lead && (lead[2].length > 60 || /[.;,]$/.test(lead[2])))
      return { text, size, align: "Justified", indent: true, after: 6, line: 1.2, boldLead: lead[1] };
    return { text, size, bold: true, align: "Justified", indent: true, after: 6, line: 1.2, keepNext: true };
  }
  const it = line.match(/^_(.+)_$/) || line.match(/^\*([^*].*[^*])\*$/);
  if (it) return { text: it[1], size, italic: true, align: "Justified", indent: true, after: 6, line: 1.2 };
  return { text: line.replace(/\*\*/g, ""), size, align: "Justified", indent: true, after: 6, line: 1.2 };
}

const empty = (size: number): Para => ({ text: "", size, align: "Centered" });
/** "Hà Nội, ngày 05 tháng 01 năm 2026" → { place, date } */
const splitPlaceDate = (t?: string) => {
  const m = String(t || "").match(/^([^,]+),\s*(ngày.*)$/iu);
  return m ? { place: m[1].trim(), date: m[2].trim() } : { place: "", date: "" };
};

/** Cụm chữ Hán, Nhật, Hàn, Ả Rập, Kirin, Thái trong các dòng (văn bản tiếng Việt không có; tên riêng nước ngoài hiếm gặp nên chỉ báo
 *  cho model sửa, không từ chối). Tối đa 6 mục dạng `<cụm chữ lạ> (sau “<chữ đứng trước>”)`; cụm chữ lạ đứng đầu để làm `find`. */
export function foreignRuns(lines: string[]): string[] {
  const out: string[] = [];
  const re = /[؀-ۿЀ-ӿ฀-๿぀-ヿ㐀-鿿가-힯]+/gu;
  for (const raw of lines) {
    const l = raw.normalize("NFC");
    for (const m of l.matchAll(re)) {
      if (out.length >= 6) return out;
      // Cắt theo ký tự (không theo đơn vị mã) để không làm gãy chữ
      const before = Array.from(l.slice(0, m.index)).slice(-20).join("").trim();
      out.push(before ? `${m[0]} (sau “${before}”)` : m[0]);
    }
  }
  return out;
}

/** So khớp hai dòng bỏ dấu, bỏ ký hiệu markdown và dấu câu (để nhận ra dòng nội dung lặp lại phần đầu hoặc phần ký) */
const key = (t: string) =>
  noAccent(clean(t))
    .toLowerCase()
    .replace(/[#*_`>]/g, "")
    .replace(/[\s.,:;()\-–—]+/g, " ")
    .trim();
/** Dòng bỏ ký hiệu markdown ở hai đầu */
const plain = (t: string) => clean(t).replace(/^[#>\s]+/, "").replace(/^[*_]+|[*_]+$/g, "").trim();

/** Giá trị bắt được từ các dòng lặp phần đầu (để không mất thông tin khi model dồn hết vào body — soát mã 4.11.8, C1) */
export interface EchoFound {
  title?: string;
  number?: string;
  placeDate?: string;
  /** Trích yếu lấy từ dòng lặp tên loại ("BIÊN BẢN ĐẠI HỘI ĐỒNG CỔ ĐÔNG …" → "Đại hội đồng cổ đông …") */
  subject?: string;
  /** Nơi nhận ở dòng "Kính gửi: …" model chép vào đầu nội dung */
  to?: string[];
  /** Số dòng của khối "Lưu ý: …" cuối hợp đồng đã bỏ */
  notesDropped?: number;
}

/** Model hay chép lại phần đầu (quốc hiệu, tiêu ngữ, tên loại, "Số: …", địa danh ngày tháng) ở đầu nội dung và phần ký, nơi nhận ở
 *  cuối nội dung, trong khi công cụ đã tự dựng những phần đó → văn bản có hai dòng "HỢP ĐỒNG …" (Chạy thật 23/11/2026). Bỏ các dòng
 *  lặp ở ĐẦU và CUỐI nội dung, chỉ khi khớp trọn dòng; dòng ở giữa nội dung không bao giờ bị bỏ. Tên, số, địa danh ngày tháng của các
 *  dòng bị bỏ được trả trong `found` để bên gọi dùng khi tham số tương ứng còn trống. */
export function stripHeaderEcho(
  lines: string[],
  o: { title?: string; org?: string; subject?: string; contract?: boolean }
): { lines: string[]; removed: number; found: EchoFound } {
  const heads = new Set([QUOC_HIEU, TIEU_NGU, o.title || "", o.org || "", o.subject || ""].map(key).filter(Boolean));
  const found: EchoFound = {};
  const isHead = (l: string) => {
    const k = key(l);
    const raw = plain(l);
    if (!k || heads.has(k) || /^[-=_~*.]{3,}$/.test(clean(l)) || /^v\/v\b/i.test(raw)) return true;
    // "Kính gửi: Sở Xây dựng …," chép lại ở đầu nội dung (công cụ đã dựng dòng Kính gửi) — thử 23/11/2026, Cohere
    const kg = raw.match(/^kính gửi\s*:?\s*(.*)$/iu);
    if (kg && raw.length <= 150) {
      const v = kg[1].replace(/[,;.]+$/, "").trim();
      if (v) found.to = [...(found.to || []), v];
      return true;
    }
    // Dòng tên loại in hoa kèm trích yếu ("BIÊN BẢN ĐẠI HỘI ĐỒNG CỔ ĐÔNG BẤT THƯỜNG NĂM 2026"): bỏ, giữ phần sau làm trích yếu
    const tk = key(o.title || "");
    if (!o.contract && tk && k.startsWith(tk + " ") && raw === upper(raw) && raw.length <= 160 && !/[;:]$/.test(raw)) {
      const rest = raw.slice(Array.from(o.title || "").length).trim();
      if (rest && !found.subject) found.subject = rest.charAt(0) + rest.slice(1).toLocaleLowerCase("vi-VN");
      return true;
    }
    // Hợp đồng: dòng tên in hoa ("HỢP ĐỒNG XÂY DỰNG"), không phải câu ("HỢP ĐỒNG NÀY ĐƯỢC LẬP …, …")
    if (o.contract && /^hop dong\b/.test(k) && raw.length <= 100 && !/[,;:.]/.test(raw) && raw === upper(raw)) {
      found.title = found.title || raw;
      return true;
    }
    // "Số: 12/HĐ-XD", "Số: …/2026/HĐTC" — phải có dấu hai chấm sau "Số" (không bắt "Số liệu …", "Sở …", "So sánh 2024/2025")
    const so = raw.match(/^số\s*:\s*(.*)$/iu);
    if (so && raw.length <= 50) {
      const v = so[1].trim();
      if (/\d/.test(v.split("/")[0])) found.number = found.number || v;
      return true;
    }
    // "Hà Nội, ngày … tháng … năm …" đứng một mình (không bắt câu "Hôm nay, ngày … năm …, chúng tôi gồm:")
    if (raw.length <= 70 && !/^hom nay\b/.test(k) && /^[^,]{1,40}, ngay\b.{0,14}\bthang\b.{0,14}\bnam\b[^a-z]*$/.test(noAccent(raw).toLowerCase())) {
      if (/\d{4}/.test(raw) && /ngày\s+\d/iu.test(raw)) found.placeDate = found.placeDate || raw;
      return true;
    }
    return false;
  };
  // Dòng nhãn phần ký: "ĐẠI DIỆN BÊN A", và với hợp đồng cả "Bên chuyển nhượng", "BÊN NHẬN CHUYỂN NHƯỢNG (Ký, ghi rõ họ tên)"
  const isSignLine = (k: string) =>
    (/^dai dien ben [ab]\b/.test(k) && k.length <= 60) ||
    (!!o.contract && k.length <= 70 && /^(dai dien )?ben (a|b|chuyen nhuong|nhan chuyen nhuong|ban|mua|cho thue|thue|giao thau|nhan thau|giao|nhan)\b/.test(k) && !/\b(co|phai|se|duoc|chiu|cam ket)\b/.test(k));
  let s = 0;
  let e = lines.length;
  while (s < e && s < 10 && isHead(lines[s])) s++;
  // Hợp đồng dán từ khung chat (thử 23/11/2026, Cohere): sau phần ký tự viết còn đoạn "Lưu ý: cần đính kèm …" — lời dặn người dùng,
  // không phải điều khoản. Có dòng ký đứng trước thì bỏ cả khối "Lưu ý" cuối để phần ký tự viết được bỏ theo quy tắc dưới.
  if (o.contract) {
    for (let i = Math.max(s, e - 12); i < e; i++) {
      if (/^(lưu ý|ghi chú)\s*:/iu.test(plain(lines[i])) && lines.slice(Math.max(s, i - 8), i).some((l) => isSignLine(key(l)))) {
        found.notesDropped = lines.slice(i, e).length;
        e = i;
        break;
      }
    }
  }
  // Phần cuối: từ dòng "ĐẠI DIỆN BÊN …" hoặc "Nơi nhận:" (trong 10 dòng cuối) tới hết, nếu sau nó chỉ là dòng ngắn của phần ký
  // (chức vụ, họ tên, "(Ký, ghi rõ họ tên…)") hoặc dòng "- …" của nơi nhận — không có đề mục, điều khoản
  for (let i = Math.max(s, e - 10); i < e; i++) {
    const k = key(lines[i]);
    const sign = isSignLine(k);
    const noiNhan = k === "noi nhan";
    if (!sign && !noiNhan) continue;
    const rest = lines.slice(i + 1, e);
    const ok = rest.every((l) => {
      const r = plain(l);
      if (/^(điều|chương|mục)\s+\d/iu.test(r) || /^#/.test(clean(l))) return false;
      return noiNhan ? /^[-–•+]/.test(clean(l)) || isSignLine(key(l)) : r.length <= 60 && !/[.;:]$/.test(r.replace(/\)$/, ""));
    });
    if (ok) {
      e = i;
      break;
    }
  }
  // Còn sót dòng "Bên A", "Bên B" đứng một mình, hoặc đường gạch "---" ở cuối nội dung
  while (e > s && ((o.contract && /^(ben [ab]|ben giao thau|ben nhan thau)$/.test(key(lines[e - 1]))) || /^[-=_~*.]{3,}$/.test(clean(lines[e - 1])))) e--;
  return { lines: lines.slice(s, e), removed: s + (lines.length - e), found };
}

/** Họ tên người ký viết thường, in đậm (thể thức): model đưa "NGUYỄN ANH QUÂN" toàn chữ hoa thì đổi thành "Nguyễn Anh Quân" */
export function personName(v: unknown): string {
  const n = clean(v);
  if (!n || n !== upper(n) || !/\p{Lu}{2}/u.test(n)) return n;
  return n
    .split(" ")
    .map((w) => (w.length <= 1 ? w : w.charAt(0) + w.slice(1).toLocaleLowerCase("vi-VN")))
    .join(" ");
}

/** Chỗ nội dung nghi sai, model phải kiểm và sửa ngay (thử model miễn phí 23/11/2026): thứ lệch ngày, từ không phải tiếng Việt */
export function contentChecks(lines: string[], refs: string[] = []): string[] {
  const out: string[] = [];
  for (const w of weekdayIssues(lines)) out.push(`Thứ trong tuần sai: ${w}`);
  const odd = strangeWords(lines);
  if (odd.length) out.push(`Từ không phải tiếng Việt hoặc hai chữ dính nhau: ${odd.join("; ")}`);
  const old = expiredLaws(lines);
  if (old.length) out.push(`Dẫn văn bản ĐÃ HẾT HIỆU LỰC: ${old.join("; ")} — sửa căn cứ hoặc để “[cần bổ sung căn cứ]”`);
  // Chữ mẫu của kỹ năng chép nguyên vào văn bản (thử 24/11/2026: Nemotron Nano, LFM ghi "THẨM QUYỀN BAN HÀNH")
  const ph = lines.filter((l) => /THẨM QUYỀN BAN HÀNH|<[^<>]{3,40}>|TÊN TỔ CHỨC|TÊN CƠ QUAN/u.test(plain(l))).map((l) => plain(l).slice(0, 60));
  if (ph.length) out.push(`Còn chữ mẫu chưa thay bằng nội dung thật: ${ph.slice(0, 3).map((x) => `“${x}”`).join("; ")}`);
  const typo = accentTypos(lines, refs);
  if (typo.length) out.push(`Viết sai dấu tên riêng hoặc cụm chuẩn: ${typo.join("; ")}`);
  return out;
}

/** Cụm chuẩn trong văn bản doanh nghiệp: so cùng chữ, khác dấu thì gần như chắc là gõ sai ("HỘI ĐÒNG", "CỔ PHẪN", "Anh Quan") */
const STD_PHRASES = ["họp Đại hội đồng cổ đông", "cuộc họp", "tiến hành họp", "nghị sự", "Hội đồng quản trị", "Đại hội đồng cổ đông", "Tổng Giám đốc", "Giám đốc", "Chủ tịch", "Chủ tọa", "cổ phần", "cổ đông", "Biên bản", "Quyết định", "Nghị quyết", "tán thành", "Ban kiểm soát", "Điều lệ"];

/** Chỗ viết cùng chữ cái nhưng khác dấu so với tên riêng người dùng đưa (tên đơn vị, người ký) và các cụm chuẩn.
 *  Chỉ so cụm từ hai chữ trở lên hoặc tên riêng; tối đa 6 mục dạng `“viết sai” → “đúng”`. */
export function accentTypos(lines: string[], refs: string[] = [], max = 6): string[] {
  const out: string[] = [];
  const phrases = [...new Set([...refs.map((r) => clean(r)).filter((r) => r.split(" ").length >= 2), ...STD_PHRASES])];
  for (const raw of lines) {
    const l = raw.normalize("NFC");
    const lo = l.toLocaleLowerCase("vi-VN");
    const base = noAccent(lo);
    if (base.length !== lo.length) continue;
    for (const ph of phrases) {
      const want = ph.normalize("NFC").toLocaleLowerCase("vi-VN");
      const key0 = noAccent(want);
      let at = base.indexOf(key0);
      while (at >= 0) {
        const before = at === 0 ? " " : base[at - 1];
        const after = base[at + key0.length] ?? " ";
        const got = lo.slice(at, at + key0.length);
        if (!/[a-z0-9]/.test(before) && !/[a-z0-9]/.test(after) && got !== want) {
          const shown = l.slice(at, at + key0.length);
          const msg = `“${shown}” → “${ph}”`;
          if (!out.includes(msg)) out.push(msg);
          if (out.length >= max) return out;
        }
        at = base.indexOf(key0, at + 1);
      }
    }
  }
  return out;
}

/** Hợp đồng (Chạy thật 23/11/2026: model dùng 'khac' cho hợp đồng → thừa tên loại, có Kính gửi, Nơi nhận, "./."):
 *  - có tên đơn vị (org): phần đầu là bảng hai cột không viền như văn bản hành chính — trái: đơn vị, "Số: …"; phải: quốc hiệu,
 *    tiêu ngữ (và địa danh, ngày nếu có);
 *  - không có org: quốc hiệu, tiêu ngữ canh giữa; "Số: …" nằm dưới tên hợp đồng;
 *  - tên hợp đồng in hoa đậm, dòng "về việc" (subject) đậm; không Kính gửi, không Nơi nhận, không "./.";
 *  - phần ký: hai cột ĐẠI DIỆN BÊN A / ĐẠI DIỆN BÊN B, họ tên hai bên cùng hàng. */
function planContract(a: AdminDocInput): AdminDocPlan {
  const size = Number(a.font_size) === 14 ? 14 : 13;
  const hdr = size - 1;
  const blanks: string[] = [];
  const notes: string[] = [];
  const blocks: Block[] = [];
  const org = clean(a.org);
  const parent = clean(a.parent_org);
  const given = upper(clean(a.title));
  const withPrefix = (t: string) => (/^HỢP ĐỒNG\b/u.test(t) ? t : `HỢP ĐỒNG ${t}`);
  const subject = clean(a.subject);
  const allLines = bodyLinesOf(a.body);
  const echo = stripHeaderEcho(allLines, { title: given ? withPrefix(given) : "", org, subject, contract: true });
  const bodyLines = echo.lines;
  // Model dồn tên, số, ngày vào body mà không đưa tham số: lấy từ các dòng vừa bỏ, không để mất (soát mã 4.11.8, C1)
  const rawTitle = given || upper(clean(echo.found.title));
  const title = rawTitle ? withPrefix(rawTitle) : TYPE_TITLE.hop_dong;
  if (!rawTitle) blanks.push("tên hợp đồng");
  if (echo.removed) notes.push(`đã bỏ ${echo.removed} dòng ở đầu hoặc cuối nội dung vì trùng phần đầu hoặc phần ký do công cụ tự dựng`);
  if (echo.found.notesDropped)
    notes.push(`đã bỏ khối “Lưu ý” ${echo.found.notesDropped} dòng ở cuối nội dung (lời dặn người dùng, không phải điều khoản): nêu lại trong câu trả lời nếu cần`);
  if (!bodyLines.length) throw new Error("Thiếu nội dung hợp đồng (body): mỗi dòng là một đoạn — căn cứ, các bên, các điều khoản.");
  if (list(a.to).length || list(a.recipients).length) notes.push("hợp đồng không có Kính gửi và Nơi nhận: đã bỏ hai phần này");

  // Số hợp đồng: giữ nguyên dạng đầy đủ người dùng đưa ("01/2026/HĐ-XD"); chỉ có số thì ghép với ký hiệu
  const rawNo = (clean(a.number) || clean(echo.found.number)).replace(/^số\s*:?\s*/i, "");
  const symbol = clean(a.symbol).replace(/^\/+/, "");
  const soHD = rawNo.includes("/") ? `Số: ${rawNo}` : `Số: ${rawNo || "     "}/${symbol || BLANK}`;
  if (!rawNo) blanks.push("số hợp đồng");
  else if (!rawNo.includes("/") && !symbol) blanks.push("ký hiệu hợp đồng");
  const pd = !clean(a.place) && !clean(a.date) ? splitPlaceDate(echo.found.placeDate) : { place: "", date: "" };
  const place = (clean(a.place) || pd.place).replace(/,$/, "");
  const date = (clean(a.date) || pd.date).replace(/^,\s*/, "");
  // Hợp đồng thường ghi ngày ký trong câu "Hôm nay, ngày …": chỉ hiện dòng địa danh, ngày ở đầu khi người dùng đưa
  const diaDanh =
    place || date ? `${place || BLANK}, ${date ? (/^ngày/i.test(date) ? date : "ngày " + date) : `ngày      tháng      năm ${a.year || BLANK}`}` : "";
  const mottoRule = (w: number) => Math.min(w - 6, estimateWidthMm(TIEU_NGU, hdr + 1, { bold: true }));

  if (org) {
    const leftW = 62;
    const rightW = TEXT_WIDTH_MM - leftW;
    const orgUp = upper(org);
    const orgRule = Math.max(12, Math.min(leftW - 8, Math.min(estimateWidthMm(orgUp, hdr, { bold: true }), leftW - 4) * 0.4));
    const leftTop: Para[] = [];
    if (parent) leftTop.push({ text: upper(parent), size: hdr, align: "Centered" });
    leftTop.push({ text: orgUp, size: hdr, bold: true, align: "Centered" }, { text: "", size: 3, align: "Centered", ruleMm: orgRule }, empty(2));
    const rightTop: Para[] = [
      { text: QUOC_HIEU, size: hdr, bold: true, align: "Centered" },
      { text: TIEU_NGU, size: hdr + 1, bold: true, align: "Centered" },
      { text: "", size: 3, align: "Centered", ruleMm: mottoRule(rightW) },
      empty(2),
    ];
    const so: Para = { text: soHD, size: 13, align: "Centered", before: 4 };
    // Không có ngày ở đầu: một hàng, "Số" cuối ô trái. Đã đo trong Word 23/11/2026: hàng hai có ô phải rỗng thì Hoàn tác để sót một đoạn rỗng
    if (diaDanh)
      blocks.push({ kind: "table", widths: [leftW, rightW], rows: [[leftTop, rightTop], [[so], [{ text: diaDanh, size: hdr + 1, italic: true, align: "Centered", before: 4 }]]] });
    else blocks.push({ kind: "table", widths: [leftW, rightW], rows: [[[...leftTop, so], rightTop]] });
  } else {
    blocks.push({ kind: "para", text: QUOC_HIEU, size: hdr, bold: true, align: "Centered" });
    blocks.push({ kind: "para", text: TIEU_NGU, size: hdr + 1, bold: true, align: "Centered" });
    blocks.push({ kind: "para", text: "", size: 3, align: "Centered", ruleMm: mottoRule(TEXT_WIDTH_MM) });
    if (diaDanh) blocks.push({ kind: "para", text: diaDanh, size: hdr + 1, italic: true, align: "Right", before: 6 });
  }

  blocks.push({ kind: "para", text: title, size: Math.max(size, 14), bold: true, align: "Centered", before: 18, line: 1.2, keepNext: true });
  if (!org) blocks.push({ kind: "para", text: soHD, size, align: "Centered", line: 1.2, keepNext: true });
  if (subject) blocks.push({ kind: "para", text: subject, size, bold: true, align: "Centered", line: 1.2, keepNext: true });
  blocks.push({ kind: "para", text: "", size: 6, align: "Centered", after: 6 });

  let preamble = true;
  for (const l of bodyLines) {
    const p = bodyPara(l, size);
    if (p.bold || /^(hôm nay|điều\s+\d|bên\s)/iu.test(p.text)) preamble = false;
    // Dòng căn cứ ở phần mở đầu in nghiêng dù model không đánh dấu _…_ (Chạy thật 23/11/2026: Nemotron gửi chữ thường);
    // câu "Căn cứ biên bản nghiệm thu, …" trong điều khoản giữ chữ đứng (soát mã 4.11.8, T3)
    if (preamble && /^căn cứ(?=\s|$)/iu.test(p.text)) p.italic = true;
    blocks.push({ kind: "para", ...p });
  }
  blocks[blocks.length - 1] = { ...(blocks[blocks.length - 1] as ParaBlock), keepNext: true, after: 12 };

  // Phần ký: cột trái Bên A, cột phải Bên B
  const side = (auth: string, ttl: string): Para[] => {
    const ps: Para[] = [{ text: auth, size, bold: true, align: "Centered", keepNext: true }];
    if (ttl) ps.push({ text: ttl, size, bold: true, align: "Centered", keepNext: true });
    ps.push({ text: "(Ký, ghi rõ họ tên, đóng dấu)", size: 12, italic: true, align: "Centered", keepNext: true });
    return ps;
  };
  const lName = personName(a.left_signer_name);
  const rName = clean(a.signer_name);
  if (!lName) blanks.push("họ tên người ký Bên A");
  if (!rName) blanks.push("họ tên người ký Bên B");
  // Nhãn "ĐẠI DIỆN BÊN A/B" do công cụ tự đặt: model đưa lại nhãn này ở chức vụ thì in hai lần, đưa lộn bên thì lộn cột
  // (thử 23/11/2026: Nemotron 3 Ultra, Nemotron 3.5 Super) → chức vụ là nhãn bên thì bỏ; quyền hạn ghi nhầm bên thì dùng mặc định
  const sideTitle = (t: unknown) => {
    const v = upper(clean(t));
    return /^ĐẠI DIỆN BÊN\b/u.test(v) || /^BÊN\s+[AB]$/u.test(v) ? "" : v;
  };
  const lAuth = upper(clean(a.left_signer_authority));
  const rAuth = upper(clean(a.signer_authority));
  const left = side(lAuth && !/BÊN\s+B$/u.test(lAuth) ? lAuth : "ĐẠI DIỆN BÊN A", sideTitle(a.left_signer_title));
  const right = side(rAuth && !/BÊN\s+A$/u.test(rAuth) ? rAuth : "ĐẠI DIỆN BÊN B", sideTitle(a.signer_title));
  // Hai cột cao bằng nhau để họ tên hai bên nằm cùng hàng
  while (left.length < right.length) left.push(empty(size));
  while (right.length < left.length) right.push(empty(size));
  for (const col of [left, right]) for (let i = 0; i < 4; i++) col.push(empty(size));
  left.push({ text: lName || BLANK, size, bold: true, align: "Centered" });
  right.push({ text: rName || BLANK, size, bold: true, align: "Centered" });
  blocks.push({ kind: "table", widths: [TEXT_WIDTH_MM / 2, TEXT_WIDTH_MM / 2], rows: [[left, right]] });
  const foreign = foreignRuns(bodyLines);
  const checks = contentChecks([...bodyLines, title, subject], [org, lName, rName]);
  return { type: "hop_dong", blocks, blanks: [...new Set(blanks)], notes, size, ...(foreign.length ? { foreign } : {}), ...(checks.length ? { checks } : {}) };
}

export function planAdminDocument(input: AdminDocInput): AdminDocPlan {
  const a = input || {};
  // "Tờ trình", "công văn", "to-trinh" đều nhận
  const rawType = noAccent(clean(a.doc_type)).toLowerCase().replace(/[\s-]+/g, "_");
  // "hợp đồng", "hop_dong_xay_dung", "hợp đồng thi công": đều là hợp đồng
  if (rawType.startsWith("hop_dong")) return planContract(a);
  let type: DocType = (DOC_TYPES as readonly string[]).includes(rawType) ? (rawType as DocType) : "khac";
  // Nghị quyết trước đây chỉ dựng được bằng khac + title: nhận lại thành loại nghị quyết; "QUYẾT ĐỊNH" kèm tên NGHỊ QUYẾT cũng vậy
  // (thử 23/11/2026: Cohere dùng quyet_dinh với title "NGHỊ QUYẾT VỀ VIỆC …")
  if ((type === "khac" || type === "quyet_dinh") && /^NGHỊ QUYẾT\b/u.test(upper(clean(a.title)))) type = "nghi_quyet";
  if (rawType.startsWith("nghi_quyet")) type = "nghi_quyet";
  const size = Number(a.font_size) === 13 ? 13 : 14;
  /** Cặp cỡ chữ Quốc hiệu / Tiêu ngữ tăng giảm cùng cỡ nội dung: 12/13 hoặc 13/14 */
  const hdr = size - 1;
  const blanks: string[] = [];
  const notes: string[] = [];
  const blocks: Block[] = [];

  const org = clean(a.org);
  if (!org) throw new Error("Thiếu tên cơ quan, tổ chức ban hành (org). Chưa biết thì hỏi người dùng, không tự đặt.");
  const parent = clean(a.parent_org);
  const bodyLines = bodyLinesOf(a.body);
  let subject = clean(a.subject).replace(/^v\/v:?\s*/i, "");
  let title = upper(clean(a.title)) || TYPE_TITLE[type];
  // Tên loại kèm trích yếu trong title ("NGHỊ QUYẾT VỀ VIỆC PHÊ DUYỆT …"): tên loại một dòng, phần sau thành trích yếu nếu chưa có
  const base = TYPE_TITLE[type];
  if (base && title !== base && title.startsWith(base + " ")) {
    const rest = title.slice(base.length).trim();
    if (!subject) subject = rest.charAt(0) + rest.slice(1).toLocaleLowerCase("vi-VN");
    title = base;
    notes.push(`tên loại chỉ ghi “${base}”; phần sau đưa xuống dòng trích yếu`);
  }
  const echo = stripHeaderEcho(bodyLines, { title, org, subject });
  if (echo.removed && !echo.lines.length) echo.found = {};
  if (echo.removed && echo.lines.length) {
    bodyLines.splice(0, bodyLines.length, ...echo.lines);
    notes.push(`đã bỏ ${echo.removed} dòng ở đầu hoặc cuối nội dung vì trùng phần đầu, Kính gửi hoặc phần ký do công cụ tự dựng`);
  }
  if (!subject && echo.found.subject) subject = echo.found.subject;
  if (!bodyLines.length) throw new Error("Thiếu nội dung văn bản (body): mỗi dòng là một đoạn.");
  if (type !== "cong_van" && !title) throw new Error(`Không nhận ra loại văn bản “${clean(a.doc_type)}”. Dùng một trong: ${DOC_TYPES.filter((t) => t !== "khac").join(", ")}; hoặc doc_type "khac" kèm title (tên loại).`);
  if (!subject) blanks.push("trích yếu nội dung");
  const checks: string[] = [];
  // Nghị quyết kết phần căn cứ bằng "QUYẾT NGHỊ:", không phải "QUYẾT ĐỊNH:" (thử 23/11/2026: Nemotron 3.5 Super)
  if (type === "nghi_quyet") {
    const k = bodyLines.findIndex((l) => /^#?\s*\**QUYẾT ĐỊNH\s*:\**$/u.test(l.trim()));
    if (k >= 0) {
      bodyLines[k] = "# QUYẾT NGHỊ:";
      notes.push("nghị quyết: đã đổi dòng “QUYẾT ĐỊNH:” thành “QUYẾT NGHỊ:”");
    }
  } else if (type !== "quyet_dinh" && type !== "khac" && bodyLines.some((l) => /^#?\s*\**QUYẾT (ĐỊNH|NGHỊ)\s*:\**$/u.test(l.trim())))
    // Tờ trình có dòng "QUYẾT ĐỊNH:" (thử 23/11/2026: Cohere) — sai loại văn bản
    checks.push(`Dòng “QUYẾT ĐỊNH:” / “QUYẾT NGHỊ:” chỉ có trong quyết định, nghị quyết; ${TYPE_TITLE[type].toLocaleLowerCase("vi-VN") || "văn bản này"} không có — xóa dòng đó`);

  // ----- Phần đầu: bảng hai cột không viền -----
  // Model hay đưa "Số 12" hoặc cả "12/CV-ABC": tách số và ký hiệu. Số đủ dạng có năm ("15/2026/NQ-HĐQT") giữ nguyên như người dùng đưa
  const echoNo = !clean(a.number) ? clean(echo.found.number) : "";
  const rawNo = (clean(a.number) || echoNo).replace(/^số\s*:?\s*/i, "").trim();
  const fullNo = /^\d+\/\d{4}\/\S+$/.test(rawNo) ? rawNo : "";
  const number = fullNo ? "" : rawNo.split("/")[0].trim();
  const symbol = fullNo ? "" : (clean(a.symbol) || (rawNo.includes("/") ? rawNo.slice(rawNo.indexOf("/") + 1) : "")).replace(/^\/+/, "");
  if (!fullNo && !number) blanks.push("số văn bản (để văn thư cấp)");
  if (!fullNo && !symbol) blanks.push("ký hiệu văn bản");
  const soKyHieu = fullNo ? `Số: ${fullNo}` : `Số: ${number || "     "}/${symbol || BLANK}`;
  const pd = !clean(a.place) && !clean(a.date) ? splitPlaceDate(echo.found.placeDate) : { place: "", date: "" };
  const place = (clean(a.place) || pd.place).replace(/,$/, "");
  if (!place && type !== "bien_ban") blanks.push("địa danh");
  let date = (clean(a.date) || pd.date).replace(/^,\s*/, "");
  if (date && !/^ngày/i.test(date)) {
    const iso = date.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    const m = iso ? [iso[0], iso[3], iso[2], iso[1]] : date.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
    if (!m && /^\d{1,2}\s+tháng/i.test(date)) date = "ngày " + date;
    if (m) date = `ngày ${m[1].padStart(2, "0")} tháng ${Number(m[2]) <= 2 ? m[2].padStart(2, "0") : String(Number(m[2]))} năm ${m[3]}`;
  }
  if (!date) {
    date = `ngày      tháng      năm ${a.year || BLANK}`;
    if (type !== "bien_ban") blanks.push("ngày tháng ban hành");
  }
  const diaDanh = `${place || BLANK}, ${date}`;
  const leftW = 62;
  const rightW = TEXT_WIDTH_MM - leftW;
  const orgUp = upper(org);
  const head: [Para[], Para[]][] = [];
  const orgRule = Math.max(12, Math.min(leftW - 8, Math.min(estimateWidthMm(orgUp, hdr, { bold: true }), leftW - 4) * 0.4));
  const mottoRule = Math.min(rightW - 6, estimateWidthMm(TIEU_NGU, hdr + 1, { bold: true }));
  const leftTop: Para[] = [];
  if (parent) leftTop.push({ text: upper(parent), size: hdr, align: "Centered" });
  leftTop.push({ text: orgUp, size: hdr, bold: true, align: "Centered" }, { text: "", size: 3, align: "Centered", ruleMm: orgRule }, empty(2));
  head.push([
    leftTop,
    [{ text: QUOC_HIEU, size: hdr, bold: true, align: "Centered" }, { text: TIEU_NGU, size: hdr + 1, bold: true, align: "Centered" }, { text: "", size: 3, align: "Centered", ruleMm: mottoRule }, empty(2)],
  ]);
  const dateCell: Para = type === "bien_ban" ? empty(size) : { text: diaDanh, size: hdr + 1, italic: true, align: "Centered", before: 4 };
  head.push([[{ text: soKyHieu, size: 13, align: "Centered", before: 4 }], [dateCell]]);
  if (type === "cong_van") head.push([[{ text: `V/v ${subject || BLANK}`, size: hdr, align: "Centered", before: 6 }], [empty(hdr)]]);
  blocks.push({ kind: "table", widths: [leftW, rightW], rows: head });

  // ----- Tên loại và trích yếu -----
  if (type !== "cong_van") {
    blocks.push({ kind: "para", text: title, size, bold: true, align: "Centered", before: 12, line: 1.2, keepNext: true });
    const trich = subject || BLANK;
    blocks.push({ kind: "para", text: trich, size, bold: true, align: "Centered", line: 1.2, keepNext: true });
    const w = Math.min(estimateWidthMm(trich, size, { bold: true }), TEXT_WIDTH_MM - 20);
    blocks.push({ kind: "para", text: "", size: 3, align: "Centered", ruleMm: Math.max(15, w * 0.4), after: 12 });
  } else {
    blocks.push({ kind: "para", text: "", size: 6, align: "Centered" });
  }

  // ----- Kính gửi -----
  let to = list(a.to);
  if (!to.length && echo.found.to?.length) to = list(echo.found.to);
  // Quyết định, nghị quyết, biên bản không có "Kính gửi" (thử 23/11/2026: Cohere đưa Kính gửi vào quyết định bổ nhiệm) → chuyển sang Nơi nhận
  const movedTo = type === "quyet_dinh" || type === "nghi_quyet" || type === "bien_ban" ? to : [];
  if (movedTo.length) {
    notes.push(`${TYPE_TITLE[type].toLocaleLowerCase("vi-VN")} không có phần Kính gửi: đã chuyển ${movedTo.length} nơi sang Nơi nhận`);
    to = [];
  }
  if (type === "cong_van" || type === "to_trinh" || type === "bao_cao") {
    if (!to.length) {
      blanks.push("nơi nhận ở phần Kính gửi");
      blocks.push({ kind: "para", text: `Kính gửi: ${BLANK}`, size, align: "Centered", after: 12, line: 1.2 });
    }
  }
  if (to.length === 1) blocks.push({ kind: "para", text: `Kính gửi: ${to[0]}`, size, align: "Centered", after: 12, line: 1.2 });
  else if (to.length > 1) {
    // Không dùng bảng ở đây: đã đo trong Word, chèn bảng ngay sau một đoạn rỗng (đường kẻ dưới trích yếu) thì bảng nuốt đoạn đó
    blocks.push({ kind: "para", text: "Kính gửi:", size, align: "Left", leftMm: 30, line: 1.2 });
    to.forEach((t, i) => blocks.push({ kind: "para", text: `- ${t}${i === to.length - 1 ? dot(t) : ";"}`, size, align: "Left", leftMm: 52, line: 1.2, after: i === to.length - 1 ? 12 : 0 }));
  }

  // ----- Nội dung -----
  const paras = bodyLines.map((l) => bodyPara(l, size));
  // Dấu kết thúc "./." chỉ thêm khi đoạn CUỐI CÙNG là lời văn thường (không phải đề mục, không kết bằng ":" hay "…")
  const lastPara = paras[paras.length - 1];
  const lastText = lastPara && !lastPara.bold && !/(:|…|\.\.\.)\s*$/.test(lastPara.text) ? lastPara : undefined;
  if (lastText && !/\.\/\.$/.test(lastText.text)) {
    // "Trân trọng," → "Trân trọng./." (không thành "Trân trọng,./." — thử 23/11/2026, Cohere)
    lastText.text = lastText.text.replace(/[.,;\s]*$/, "") + "./.";
    notes.push("đã thêm dấu kết thúc “./.” vào cuối nội dung");
  }
  for (const p of paras) blocks.push({ kind: "para", ...p });
  blocks[blocks.length - 1] = { ...(blocks[blocks.length - 1] as ParaBlock), keepNext: true, after: 12 };

  // ----- Phần ký và nơi nhận -----
  const sig = normalizeSigner(a.signer_authority, a.signer_title);
  const authority = sig.authority;
  const sTitle = sig.title;
  notes.push(...sig.notes);
  checks.push(...sig.checks);
  if (type === "nghi_quyet" && !/^TM\. /u.test(authority))
    checks.push("Nghị quyết là của tập thể: dòng quyền hạn ký phải là “TM. <tên tập thể>” (ví dụ “TM. HỘI ĐỒNG QUẢN TRỊ”), chức vụ bên dưới là “CHỦ TỊCH”; sửa đoạn phần ký bằng word_set_paragraph_text hoặc báo người dùng");
  if (type === "to_trinh" && /^TM\. /u.test(authority) && to.some((t) => key(t).includes(key(authority.slice(4)))))
    checks.push(`Tờ trình gửi ${to.join(", ")} mà người trình lại ký “${authority}” (thay mặt chính nơi nhận) — người trình ký bằng chức vụ của mình, bỏ quyền hạn TM.`);
  // Bỏ học hàm, học vị ở đầu tên (không bỏ "Ông", "Bà": có thể là họ)
  const sName = personName(a.signer_name).replace(/^(?:(?:pgs|gs|ths|ts|ks|ls)\.?,?\s*|(?:luật sư|tiến sĩ|thạc sĩ|giáo sư|phó giáo sư|kỹ sư)\s+)+/i, "");
  if (!sTitle && !authority) blanks.push("quyền hạn, chức vụ người ký");
  if (!sName) blanks.push("họ tên người ký");
  const lTitle = upper(clean(a.left_signer_title));
  const lName = personName(a.left_signer_name);
  let recipients = [...movedTo, ...list(a.recipients).filter((r) => !movedTo.includes(r))];
  const needsNhuTren = to.length > 0 || type === "cong_van" || type === "to_trinh" || type === "bao_cao";
  if (needsNhuTren && !recipients.some((r) => /^như trên/i.test(r))) recipients.unshift("Như trên");
  const isLuu = (r: string) => /^lưu(\s*:|\s+vt)/i.test(r);
  recipients = recipients.map((r) => (isLuu(r) ? r.replace(/^lưu\s*:?\s*/i, "Lưu: ") : r));
  if (!recipients.some(isLuu)) {
    recipients.push(type === "bien_ban" ? "Lưu: VT, Hồ sơ" : `Lưu: VT, ${BLANK}`);
    if (type !== "bien_ban") blanks.push("đơn vị soạn thảo ở dòng Lưu");
  }
  recipients = [...recipients.filter((r) => !isLuu(r)), ...recipients.filter(isLuu).slice(0, 1)];
  const right: Para[] = [];
  if (authority) right.push({ text: authority, size, bold: true, align: "Centered" });
  right.push({ text: sTitle || (authority ? "" : BLANK), size, bold: true, align: "Centered" });
  const gap = 4;
  for (let i = 0; i < gap; i++) right.push(empty(size));
  right.push({ text: sName || BLANK, size, bold: true, align: "Centered" });
  const rightClean = right.filter((p, i) => !(i === 1 && authority && !p.text));
  const left: Para[] = [];
  if (lTitle) {
    left.push({ text: lTitle, size, bold: true, align: "Centered" });
    for (let i = 0; i < rightClean.length - 2; i++) left.push(empty(size));
    left.push({ text: lName || BLANK, size, bold: true, align: "Centered" });
    if (!lName) blanks.push("họ tên người ký bên trái");
  } else {
    left.push({ text: "Nơi nhận:", size: 12, bold: true, italic: true, align: "Left" });
    recipients.forEach((r, i) => left.push({ text: `- ${r}${i === recipients.length - 1 ? dot(r) : ";"}`, size: 11, align: "Left" }));
  }
  blocks.push({ kind: "table", widths: lTitle ? [TEXT_WIDTH_MM / 2, TEXT_WIDTH_MM / 2] : [78, TEXT_WIDTH_MM - 78], rows: [[left, rightClean]] });
  if (lTitle) {
    blocks.push({ kind: "para", text: "Nơi nhận:", size: 12, bold: true, italic: true, align: "Left", before: 12 });
    recipients.forEach((r, i) => blocks.push({ kind: "para", text: `- ${r}${i === recipients.length - 1 ? dot(r) : ";"}`, size: 11, align: "Left" }));
  }
  const foreign = foreignRuns(bodyLines);
  checks.push(...contentChecks([...bodyLines, title, subject, authority, sTitle, lTitle], [org, sName, lName]));
  return { type, blocks, blanks: [...new Set(blanks)], notes, size, ...(foreign.length ? { foreign } : {}), ...(checks.length ? { checks } : {}) };
}

/** Quyền hạn ký chỉ có các dạng TM., KT., TL., TUQ., Q. kèm tên tập thể hoặc chức vụ người đứng đầu (thử 23/11/2026, model miễn phí
 *  đưa "TGĐ", "ĐẠI DIỆN PHÁP LUẬT", "TM." trơn, "KÝ THAY TỔNG GIÁM ĐỐC", "CHỦ TỊCH HỘI ĐỒNG QUẢN TRỊ" trùng chức vụ):
 *  viết đủ chữ thì đổi ra dạng tắt; trùng chức vụ thì bỏ; sai dạng thì bỏ và báo. */
export function normalizeSigner(rawAuth: unknown, rawTitle: unknown): { authority: string; title: string; notes: string[]; checks: string[] } {
  const notes: string[] = [];
  const checks: string[] = [];
  let auth = upper(clean(rawAuth)).replace(/\s+/g, " ");
  let title = upper(clean(rawTitle)).replace(/\s+/g, " ");
  const full: [RegExp, string][] = [
    [/^THAY MẶT\s+/u, "TM. "],
    [/^KÝ THAY\s+/u, "KT. "],
    [/^THỪA ỦY QUYỀN\s+|^THỪA UỶ QUYỀN\s+/u, "TUQ. "],
    [/^THỪA LỆNH\s+/u, "TL. "],
  ];
  for (const [re, abbr] of full) if (re.test(auth)) auth = auth.replace(re, abbr);
  auth = auth.replace(/^(TM|KT|TL|TUQ|Q)\s*\.?\s+/u, "$1. ");
  const valid = /^(TM|KT|TL|TUQ|Q)\. \S/u.test(auth);
  if (auth && !valid) {
    // "TM." trơn, chữ viết tắt chức vụ, trùng chức vụ, hay nhãn tự đặt: không phải quyền hạn ký
    if (!title && auth.length > 3 && !/^(TM|KT|TL|TUQ|Q)\.?$/u.test(auth) && /^(PHÓ\s+)?(TỔNG\s+)?(GIÁM ĐỐC|CHỦ TỊCH|TRƯỞNG|THƯ KÝ|CHỦ TỌA|KẾ TOÁN TRƯỞNG)(?=\s|$)/u.test(auth)) {
      title = auth;
      notes.push(`“${auth}” là chức vụ, không phải quyền hạn ký: đã đưa xuống dòng chức vụ`);
    } else if (auth !== title) {
      notes.push(`đã bỏ “${auth}” ở dòng quyền hạn ký (chỉ dùng TM., KT., TL., TUQ., Q. kèm tên tập thể hoặc chức vụ người đứng đầu)`);
    }
    auth = "";
  }
  if (auth && title && auth === title) title = "";
  // "KT. PHÓ TỔNG GIÁM ĐỐC" + "PHÓ TỔNG GIÁM ĐỐC": sau KT. phải là chức vụ người đứng đầu, dòng dưới mới là cấp phó
  // (\b của JavaScript không nhận chữ có dấu: "PHÓ" phải so bằng khoảng trắng phía sau)
  if (/^KT\. /u.test(auth) && /^PHÓ\s/u.test(auth.slice(4)))
    checks.push(`Quyền hạn ký “${auth}” sai: sau KT. là chức vụ người đứng đầu (ví dụ “KT. TỔNG GIÁM ĐỐC”), dòng chức vụ bên dưới mới là cấp phó`);
  return { authority: auth, title, notes, checks };
}

/** Một dòng tóm tắt hiện trên khung */
export function adminDocBrief(plan: AdminDocPlan): string {
  const name = plan.type === "cong_van" ? "công văn" : (TYPE_TITLE[plan.type] || "văn bản").toLocaleLowerCase("vi-VN");
  return `Đã dựng ${name} đúng thể thức${plan.blanks.length ? ` · còn ${plan.blanks.length} chỗ cần điền` : ""}${plan.checks?.length ? ` · ${plan.checks.length} chỗ nghi sai cần kiểm` : ""}`;
}
