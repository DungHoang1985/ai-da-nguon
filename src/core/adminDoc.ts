/** Dựng bố cục văn bản hành chính (công văn, tờ trình, báo cáo, quyết định, thông báo, giấy mời, biên bản…) đúng thể thức:
 *  hàm thuần, biến các trường AI đưa thành một "bản vẽ" gồm bảng không viền và các đoạn đã có sẵn cỡ chữ, kiểu chữ, căn lề.
 *  Phần Office.js (wordAdmin.ts) chỉ việc dựng theo bản vẽ, nên thể thức đúng dù model mạnh hay yếu.
 *  Số đo theo chuẩn trình bày văn bản hành chính đang áp dụng: A4, Times New Roman, cỡ chữ 13–14. */

export const DOC_TYPES = ["cong_van", "to_trinh", "bao_cao", "quyet_dinh", "thong_bao", "giay_moi", "bien_ban", "ke_hoach", "hop_dong", "khac"] as const;
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

/** Cụm chữ Hán, Nhật, Hàn, Ả Rập, Kirin, Thái trong các dòng nội dung (văn bản tiếng Việt không có; tên riêng nước ngoài hiếm gặp
 *  nên chỉ báo cho model sửa, không từ chối). Trả tối đa 6 cụm, mỗi cụm kèm chữ đứng trước để tìm được chỗ. */
export function foreignRuns(lines: string[]): string[] {
  const out: string[] = [];
  const re = /[\u0600-\u06FF\u0400-\u04FF\u0E00-\u0E7F\u3040-\u30FF\u3400-\u9FFF\uAC00-\uD7AF]+/gu;
  for (const l of lines)
    for (const m of l.matchAll(re)) {
      if (out.length >= 6) return out;
      out.push(`${l.slice(Math.max(0, (m.index || 0) - 20), m.index).trimStart()}${m[0]}`);
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

/** Model hay chép lại phần đầu (quốc hiệu, tiêu ngữ, tên loại, "Số: …", địa danh ngày tháng) ở đầu nội dung và phần ký ở cuối
 *  nội dung, trong khi công cụ đã tự dựng những phần đó → văn bản có hai dòng "HỢP ĐỒNG …" (Chạy thật 23/11/2026). Bỏ các dòng
 *  lặp ở ĐẦU và CUỐI nội dung, chỉ khi khớp trọn dòng; dòng ở giữa nội dung không bao giờ bị bỏ. */
export function stripHeaderEcho(
  lines: string[],
  o: { title?: string; org?: string; subject?: string; contract?: boolean }
): { lines: string[]; removed: number } {
  const heads = new Set([QUOC_HIEU, TIEU_NGU, o.title || "", o.org || "", o.subject || ""].map(key).filter(Boolean));
  const isHead = (l: string) => {
    const k = key(l);
    const raw = clean(l);
    return (
      !k ||
      heads.has(k) ||
      // Hợp đồng: dòng tên hợp đồng in hoa khác chữ với title ("HỢP ĐỒNG XÂY DỰNG" khi title là "HỢP ĐỒNG THI CÔNG XÂY DỰNG")
      (!!o.contract && /^hop dong\b/.test(k) && raw.length <= 100 && raw.replace(/[*_#]/g, "") === upper(raw.replace(/[*_#]/g, ""))) ||
      /^[-=_~*.]{3,}$/.test(raw) ||
      // "Số: 12/HĐ-XD", "Số: …/2026/HĐTC" — dòng ngắn bắt đầu bằng "Số" và có dấu "/"
      (/^so\b/.test(k) && raw.length <= 40 && raw.includes("/")) ||
      /^v\/v\b/i.test(raw) ||
      // "Hà Nội, ngày … tháng … năm …" đứng một mình (không bắt câu "Hôm nay, ngày … năm …, chúng tôi gồm:")
      (raw.length <= 70 && !/^hom nay\b/.test(k) && /^[^,]{1,40}, ngay\b.{0,14}\bthang\b.{0,14}\bnam\b[^a-z]*$/.test(noAccent(raw).toLowerCase()))
    );
  };
  const isFoot = (l: string) => {
    const k = key(l);
    return (
      !k ||
      k === "noi nhan" ||
      (/^dai dien ben [ab]\b/.test(k) && k.length <= 60) ||
      /^ky,? (ghi ro )?ho ten/.test(k) ||
      (!!o.contract && /^(ben [ab]|ben giao thau|ben nhan thau)$/.test(k))
    );
  };
  let s = 0;
  let e = lines.length;
  while (s < e && s < 10 && isHead(lines[s])) s++;
  let tail = 0;
  while (e > s && tail < 8 && isFoot(lines[e - 1])) {
    e--;
    tail++;
  }
  return { lines: lines.slice(s, e), removed: s + (lines.length - e) };
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
  const rawTitle = upper(clean(a.title));
  const title = rawTitle ? (/^HỢP ĐỒNG\b/u.test(rawTitle) ? rawTitle : `HỢP ĐỒNG ${rawTitle}`) : TYPE_TITLE.hop_dong;
  if (!rawTitle) blanks.push("tên hợp đồng");
  const subject = clean(a.subject);
  const allLines = (Array.isArray(a.body) ? (a.body as unknown[]).map((x) => String(x ?? "")).join("\n") : String(a.body ?? ""))
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  const echo = stripHeaderEcho(allLines, { title, org, subject, contract: true });
  const bodyLines = echo.lines;
  if (echo.removed) notes.push(`đã bỏ ${echo.removed} dòng ở đầu hoặc cuối nội dung vì trùng phần đầu hoặc phần ký do công cụ tự dựng`);
  if (!bodyLines.length) throw new Error("Thiếu nội dung hợp đồng (body): mỗi dòng là một đoạn — căn cứ, các bên, các điều khoản.");
  if (list(a.to).length || list(a.recipients).length) notes.push("hợp đồng không có Kính gửi và Nơi nhận: đã bỏ hai phần này");

  // Số hợp đồng: giữ nguyên dạng đầy đủ người dùng đưa ("01/2026/HĐ-XD"); chỉ có số thì ghép với ký hiệu
  const rawNo = clean(a.number).replace(/^số\s*:?\s*/i, "");
  const symbol = clean(a.symbol).replace(/^\/+/, "");
  const soHD = rawNo.includes("/") ? `Số: ${rawNo}` : `Số: ${rawNo || "     "}/${symbol || BLANK}`;
  if (!rawNo) blanks.push("số hợp đồng");
  else if (!rawNo.includes("/") && !symbol) blanks.push("ký hiệu hợp đồng");
  const place = clean(a.place).replace(/,$/, "");
  const date = clean(a.date).replace(/^,\s*/, "");
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

  for (const l of bodyLines) {
    const p = bodyPara(l, size);
    // Dòng căn cứ của hợp đồng in nghiêng dù model không đánh dấu _…_ (Chạy thật 23/11/2026: Nemotron gửi chữ thường)
    if (!p.bold && /^căn cứ(?=\s|$)/iu.test(p.text)) p.italic = true;
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
  const lName = clean(a.left_signer_name);
  const rName = clean(a.signer_name);
  if (!lName) blanks.push("họ tên người ký Bên A");
  if (!rName) blanks.push("họ tên người ký Bên B");
  const left = side(upper(clean(a.left_signer_authority)) || "ĐẠI DIỆN BÊN A", upper(clean(a.left_signer_title)));
  const right = side(upper(clean(a.signer_authority)) || "ĐẠI DIỆN BÊN B", upper(clean(a.signer_title)));
  // Hai cột cao bằng nhau để họ tên hai bên nằm cùng hàng
  while (left.length < right.length) left.push(empty(size));
  while (right.length < left.length) right.push(empty(size));
  for (const col of [left, right]) for (let i = 0; i < 4; i++) col.push(empty(size));
  left.push({ text: lName || BLANK, size, bold: true, align: "Centered" });
  right.push({ text: rName || BLANK, size, bold: true, align: "Centered" });
  blocks.push({ kind: "table", widths: [TEXT_WIDTH_MM / 2, TEXT_WIDTH_MM / 2], rows: [[left, right]] });
  const foreign = foreignRuns(bodyLines);
  return { type: "hop_dong", blocks, blanks: [...new Set(blanks)], notes, size, ...(foreign.length ? { foreign } : {}) };
}

export function planAdminDocument(input: AdminDocInput): AdminDocPlan {
  const a = input || {};
  // "Tờ trình", "công văn", "to-trinh" đều nhận
  const rawType = noAccent(clean(a.doc_type)).toLowerCase().replace(/[\s-]+/g, "_");
  // "hợp đồng", "hop_dong_xay_dung", "hợp đồng thi công": đều là hợp đồng
  if (rawType.startsWith("hop_dong")) return planContract(a);
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
  const subject = clean(a.subject).replace(/^v\/v:?\s*/i, "");
  const title = upper(clean(a.title)) || TYPE_TITLE[type];
  const echo = stripHeaderEcho(bodyLines, { title, org, subject });
  if (echo.removed && echo.lines.length) {
    bodyLines.splice(0, bodyLines.length, ...echo.lines);
    notes.push(`đã bỏ ${echo.removed} dòng ở đầu hoặc cuối nội dung vì trùng phần đầu hoặc phần ký do công cụ tự dựng`);
  }
  if (!bodyLines.length) throw new Error("Thiếu nội dung văn bản (body): mỗi dòng là một đoạn.");
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
  const foreign = foreignRuns(bodyLines);
  return { type, blocks, blanks: [...new Set(blanks)], notes, size, ...(foreign.length ? { foreign } : {}) };
}

/** Một dòng tóm tắt hiện trên khung */
export function adminDocBrief(plan: AdminDocPlan): string {
  const name = plan.type === "cong_van" ? "công văn" : (TYPE_TITLE[plan.type] || "văn bản").toLocaleLowerCase("vi-VN");
  return `Đã dựng ${name} đúng thể thức${plan.blanks.length ? ` · còn ${plan.blanks.length} chỗ cần điền` : ""}`;
}
