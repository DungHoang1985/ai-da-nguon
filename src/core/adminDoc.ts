/** Dựng bố cục văn bản hành chính (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản…) đúng thể thức:
 *  hàm thuần, biến các trường AI đưa thành một "bản vẽ" gồm bảng không viền và các đoạn đã có sẵn cỡ chữ, kiểu chữ, căn lề.
 *  Phần Office.js (wordAdmin.ts) chỉ việc dựng theo bản vẽ, nên thể thức đúng dù model mạnh hay yếu.
 *  Số đo theo chuẩn trình bày văn bản hành chính đang áp dụng: A4, Times New Roman, cỡ chữ 13–14. */

export const DOC_TYPES = ["cong_van", "to_trinh", "bao_cao", "quyet_dinh", "thong_bao", "giay_moi", "bien_ban", "ke_hoach", "khac"] as const;
export type DocType = (typeof DOC_TYPES)[number];

/** Tên loại in trên văn bản (công văn không có tên loại) */
export const TYPE_TITLE: Record<DocType, string> = {
  cong_van: "",
  to_trinh: "TỜ TRÌNH",
  bao_cao: "BÁO CÁO",
  quyet_dinh: "QUYẾT ĐỊNH",
  thong_bao: "THÔNG BÁO",
  giay_moi: "GIẤY MỜI",
  bien_ban: "BIÊN BẢN",
  ke_hoach: "KẾ HOẠCH",
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
}

const clean = (s: unknown) =>
  String(s ?? "")
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
        .replace(/(\p{L}{4,})\.+$/u, "$1")
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

/** Tách một dòng nội dung thành đoạn có định dạng */
function bodyPara(raw: string, size: number): Para {
  const line = raw.trim();
  const h1 = line.match(/^#\s+(.*)$/);
  if (h1) return { text: h1[1].replace(/\*\*/g, ""), size, bold: true, align: "Centered", after: 6, before: 6, line: 1.2, keepNext: true };
  const h2 = line.match(/^#{2,4}\s+(.*)$/) || line.match(/^\*\*(.+)\*\*$/);
  if (h2) return { text: h2[1].replace(/\*\*/g, ""), size, bold: true, align: "Justified", indent: true, after: 6, line: 1.2, keepNext: true };
  const it = line.match(/^_(.+)_$/) || line.match(/^\*([^*].*[^*])\*$/);
  if (it) return { text: it[1], size, italic: true, align: "Justified", indent: true, after: 6, line: 1.2 };
  return { text: line.replace(/\*\*/g, ""), size, align: "Justified", indent: true, after: 6, line: 1.2 };
}

const empty = (size: number): Para => ({ text: "", size, align: "Centered" });

export function planAdminDocument(input: AdminDocInput): AdminDocPlan {
  const a = input || {};
  // "Tờ trình", "công văn", "to-trinh" đều nhận
  const rawType = noAccent(clean(a.doc_type)).toLowerCase().replace(/[\s-]+/g, "_");
  const type: DocType = (DOC_TYPES as readonly string[]).includes(rawType) ? (rawType as DocType) : "khac";
  const size = Number(a.font_size) === 13 ? 13 : 14;
  /** Cặp cỡ chữ Quốc hiệu / Tiêu ngữ tăng giảm cùng cỡ nội dung: 12/13 hoặc 13/14 */
  const hdr = size - 1;
  const blanks: string[] = [];
  const notes: string[] = [];
  const blocks: Block[] = [];

  const org = clean(a.org);
  if (!org) throw new Error("Thiếu tên cơ quan, tổ chức ban hành (org). Chưa biết thì hỏi người dùng, không tự đặt.");
  const parent = clean(a.parent_org);
  const bodyLines = (Array.isArray(a.body) ? (a.body as unknown[]).map((x) => String(x ?? "")).join("\n") : String(a.body ?? ""))
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!bodyLines.length) throw new Error("Thiếu nội dung văn bản (body): mỗi dòng là một đoạn.");
  const subject = clean(a.subject).replace(/^v\/v:?\s*/i, "");
  const title = upper(clean(a.title)) || TYPE_TITLE[type];
  if (type !== "cong_van" && !title) throw new Error(`Không nhận ra loại văn bản “${clean(a.doc_type)}”. Dùng một trong: ${DOC_TYPES.filter((t) => t !== "khac").join(", ")}; hoặc doc_type "khac" kèm title (tên loại).`);
  if (!subject) blanks.push("trích yếu nội dung");

  // ----- Phần đầu: bảng hai cột không viền -----
  // Model hay đưa "Số 12" hoặc cả "12/CV-ABC": chỉ lấy phần số
  const number = clean(a.number)
    .replace(/^số\s*:?\s*/i, "")
    .split("/")[0]
    .trim();
  const symbol = clean(a.symbol).replace(/^\/+/, "");
  if (!number) blanks.push("số văn bản (để văn thư cấp)");
  if (!symbol) blanks.push("ký hiệu văn bản");
  const soKyHieu = `Số: ${number || "     "}/${symbol || BLANK}`;
  const place = clean(a.place).replace(/,$/, "");
  if (!place && type !== "bien_ban") blanks.push("địa danh");
  let date = clean(a.date).replace(/^,\s*/, "");
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
  const to = list(a.to);
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
    lastText.text = lastText.text.replace(/[.\s]*$/, "") + "./.";
    notes.push("đã thêm dấu kết thúc “./.” vào cuối nội dung");
  }
  for (const p of paras) blocks.push({ kind: "para", ...p });
  blocks[blocks.length - 1] = { ...(blocks[blocks.length - 1] as ParaBlock), keepNext: true, after: 12 };

  // ----- Phần ký và nơi nhận -----
  const authority = upper(clean(a.signer_authority));
  const sTitle = upper(clean(a.signer_title));
  // Bỏ học hàm, học vị ở đầu tên (không bỏ "Ông", "Bà": có thể là họ)
  const sName = clean(a.signer_name).replace(/^(?:(?:pgs|gs|ths|ts|ks|ls)\.?,?\s*|(?:luật sư|tiến sĩ|thạc sĩ|giáo sư|phó giáo sư|kỹ sư)\s+)+/i, "");
  if (!sTitle && !authority) blanks.push("quyền hạn, chức vụ người ký");
  if (!sName) blanks.push("họ tên người ký");
  const lTitle = upper(clean(a.left_signer_title));
  const lName = clean(a.left_signer_name);
  let recipients = list(a.recipients);
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
  return { type, blocks, blanks: [...new Set(blanks)], notes, size };
}

/** Một dòng tóm tắt hiện trên khung */
export function adminDocBrief(plan: AdminDocPlan): string {
  const name = plan.type === "cong_van" ? "công văn" : (TYPE_TITLE[plan.type] || "văn bản").toLocaleLowerCase("vi-VN");
  return `Đã dựng ${name} đúng thể thức${plan.blanks.length ? ` · còn ${plan.blanks.length} chỗ cần điền` : ""}`;
}
