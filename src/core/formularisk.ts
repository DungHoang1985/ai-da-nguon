/** Nhận diện công thức Excel cần hỏi người dùng trước khi ghi (Đợt 4): hàm gọi ra mạng, chạy mã ngoài, tham chiếu gián tiếp,
 *  liên kết tới tệp khác. Hàm thuần. Ưu tiên không báo nhầm: tên hàm nằm trong chuỗi ("…") hay trong tên sheet ('…') không tính,
 *  tham chiếu bảng kiểu Bảng1[Cột] không phải liên kết ngoài. */

export interface FormulaRisk {
  /** Tên hàm hoặc loại, ví dụ "WEBSERVICE", "liên kết tệp ngoài" */
  what: string;
  why: string;
  formula: string;
  /** Vị trí trong khối giá trị: hàng, cột tính từ 0 */
  row: number;
  col: number;
}

const FUNCS: { re: RegExp; what: string; why: string }[] = [
  { re: /\bWEBSERVICE\s*\(/i, what: "WEBSERVICE", why: "gửi yêu cầu ra mạng mỗi lần tính lại, có thể mang dữ liệu trong sổ tính ra ngoài" },
  { re: /\bIMAGE\s*\(/i, what: "IMAGE", why: "tải ảnh từ một địa chỉ mạng mỗi lần mở tệp" },
  { re: /\bRTD\s*\(/i, what: "RTD", why: "gọi một chương trình ngoài Excel để lấy dữ liệu" },
  { re: /\b(CALL|REGISTER(\.ID)?|EXEC)\s*\(/i, what: "CALL / REGISTER / EXEC", why: "chạy mã hoặc chương trình ngoài Excel" },
  { re: /\bFILTERXML\s*\(/i, what: "FILTERXML", why: "thường đi cùng WEBSERVICE để đọc dữ liệu lấy từ mạng" },
  { re: /\bINDIRECT\s*\(/i, what: "INDIRECT", why: "tham chiếu gián tiếp: không lần được nguồn số liệu, tính lại liên tục, hỏng khi đổi tên sheet" },
];
/** Liên kết bấm được dựng từ nội dung ô (ghép chuỗi) hoặc trỏ ra mạng: địa chỉ có thể mang dữ liệu đi khi người dùng bấm */
const HYPERLINK = /\bHYPERLINK\s*\(/i;
/** Kiểu DDE cũ: =cmd|'/c calc'!A1, =1+cmd|'/c calc'!A1, ='cmd'|'…'!A1. Excel không có toán tử "|": sau khi bỏ chuỗi "…",
 *  dấu | theo sau là tham số trong nháy đơn hoặc tên rồi dấu ! chỉ có thể là DDE. */
const DDE = /\|\s*(?:'[^']*'|[^'!|\s]+)\s*!/;
/** Tham chiếu tệp khác: [Tệp.xlsx]Sheet!A1, 'C:\thư mục\[Tệp.xlsx]Sheet'!A1, [1]Sheet!A1, hoặc đường dẫn mạng */
const EXTERNAL_IN_QUOTES = /'[^']*\[[^\]']+\][^']*'\s*!/;
/** [Tệp]Sheet!…: ngay sau ngoặc vuông là tên sheet rồi dấu ! (tham chiếu bảng kiểu [2024], [@Cột] không có tên sheet liền sau) */
const EXTERNAL_BARE = /(^|[=+\-*/^&(,;<> ])\[[^\]]+\][\p{L}\p{N}_ .]+!/iu;
/** 'C:\thư mục\Tệp.xlsx'!Tên hoặc Tệp.xlsx!Tên (tên vùng của tệp khác) */
const EXTERNAL_FILE_NAME = /(?:'[^']*\.(?:xl[a-z]{1,2}|csv|ods)'|[\p{L}\p{N}_]+\.(?:xls[xmb]?|csv))\s*!/iu;
const NET_PATH = /'(?:https?:\/\/|\\\\)[^']*'\s*!/i;

/** Bỏ chuỗi "…" (kể cả "" là dấu nháy kép trong chuỗi) để tên hàm trong chữ không bị tính */
export function stripStrings(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, '""');
}
/** Bỏ tên sheet đặt trong nháy đơn ('Kế hoạch INDIRECT (cũ)'!A1) */
const stripSheetNames = (f: string) => f.replace(/'(?:[^']|'')*'(?=\s*!)/g, "S");

export function formulaRisksOf(formula: string): { what: string; why: string }[] {
  const f = String(formula ?? "");
  // Chữ bắt đầu bằng + - @ chỉ được coi là công thức khi ký tự kế tiếp không phải khoảng trắng ("- Họp call (9h)" là gạch đầu dòng)
  if (!f.startsWith("=") && !/^[+\-@]\S/.test(f)) return [];
  const out: { what: string; why: string }[] = [];
  const noStr = stripStrings(f);
  const code = stripSheetNames(noStr);
  for (const x of FUNCS) if (x.re.test(code)) out.push({ what: x.what, why: x.why });
  // Liên kết nội bộ ("#Sheet!A1", kể cả ghép với tên sheet) không tính; chỉ hỏi khi trỏ ra ngoài hoặc địa chỉ ghép từ nội dung ô
  const internalLink = /\bHYPERLINK\s*\(\s*"#/i.test(f);
  if (HYPERLINK.test(code) && !internalLink && (/&|\bCONCAT(ENATE)?\s*\(|\bTEXTJOIN\s*\(|\bSUBSTITUTE\s*\(|\bREPLACE\s*\(/i.test(code) || /"\s*(https?:|ftp:|file:|\\\\)/i.test(f)))
    out.push({ what: "HYPERLINK", why: "tạo liên kết bấm ra ngoài; địa chỉ ghép từ nội dung ô có thể mang dữ liệu đi khi bấm" });
  if (DDE.test(code)) out.push({ what: "liên kết DDE", why: "kiểu liên kết cũ có thể chạy chương trình trên máy" });
  if (EXTERNAL_IN_QUOTES.test(noStr) || EXTERNAL_BARE.test(code) || EXTERNAL_FILE_NAME.test(noStr) || NET_PATH.test(noStr))
    out.push({ what: "liên kết tệp ngoài", why: "số liệu lấy từ tệp khác: người nhận tệp không có tệp đó sẽ thấy số cũ hoặc lỗi" });
  return out;
}

/** Quét một khối giá trị sắp ghi (mảng hai chiều; phần tử không phải chuỗi thì bỏ qua). */
export function findFormulaRisks(values: any): FormulaRisk[] {
  const rows: any[] = Array.isArray(values) ? values : [[values]];
  const out: FormulaRisk[] = [];
  rows.forEach((r, i) => {
    const row: any[] = Array.isArray(r) ? r : [r];
    row.forEach((v, j) => {
      if (typeof v !== "string") return;
      for (const x of formulaRisksOf(v.trim())) out.push({ ...x, formula: v.trim(), row: i, col: j });
    });
  });
  return out;
}

/** Câu cảnh báo cho hộp hỏi; null = không có gì phải hỏi. addr đổi vị trí trong khối thành địa chỉ ô (nếu biết ô đầu). */
export function formulaRiskMessage(values: any, addr?: (row: number, col: number) => string): string | null {
  const risks = findFormulaRisks(values);
  if (!risks.length) return null;
  const byWhat = new Map<string, FormulaRisk[]>();
  for (const r of risks) byWhat.set(r.what, [...(byWhat.get(r.what) || []), r]);
  const parts = [...byWhat.entries()].map(([what, list]) => {
    const first = list[0];
    const where = addr ? ` ở ${addr(first.row, first.col)}${list.length > 1 ? ` và ${list.length - 1} ô khác` : ""}` : list.length > 1 ? ` (${list.length} ô)` : "";
    const f = first.formula.length > 140 ? first.formula.slice(0, 139) + "…" : first.formula;
    return `${what}${where} — ${first.why}. Công thức: ${f}`;
  });
  return `Công thức sắp ghi có hàm cần anh xem trước: ${parts.join(" · ")}`;
}
