/** Dấu vết soát lại cho Excel (Đợt 4): dựng các dòng của sheet "Nhật ký AI" từ một thao tác ghi. Hàm thuần, kiểm thử được ngoài Office.
 *  Nhật ký là việc của add-in: không vào lịch sử gửi model, không đổi lời nhắc hệ thống. */

export const LOG_SHEET = "Nhật ký AI";
export const LOG_HEADER = ["Thời gian", "Sheet", "Ô / vùng", "Giá trị cũ", "Giá trị mới", "Thao tác và lý do", "Yêu cầu của lượt"];
/** Màu nền tô ô AI vừa ghi (vàng nhạt) */
export const HIGHLIGHT_COLOR = "#FFF2CC";
/** Một thao tác ghi nhiều ô: ghi riêng từng ô tới mức này, phần còn lại gộp một dòng */
export const MAX_CELL_ROWS = 20;

export interface CellChange {
  /** Địa chỉ ô không kèm sheet, ví dụ "B4" */
  cell: string;
  before: any;
  after: any;
}

export interface LogInput {
  time: Date;
  sheet: string;
  /** Vùng thao tác (không kèm sheet) */
  address: string;
  /** Các ô đổi nội dung (nếu công cụ biết); không có = thao tác không đổi nội dung ô (định dạng, biểu đồ, thiết lập in…) */
  changes?: CellChange[];
  /** Việc đã làm, ví dụ "Ghi dữ liệu vào Sheet1!B2" */
  action: string;
  /** Lời AI nói ở bước đó (lý do) */
  reason?: string;
  /** Yêu cầu của người dùng ở lượt đó */
  request?: string;
}

const two = (n: number) => String(n).padStart(2, "0");

/** "16/11/2026 16:45:03" theo giờ máy */
export function formatLogTime(d: Date): string {
  return `${two(d.getDate())}/${two(d.getMonth() + 1)}/${d.getFullYear()} ${two(d.getHours())}:${two(d.getMinutes())}:${two(d.getSeconds())}`;
}

const oneLine = (s: string, max: number) => {
  const t = String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
  return t.length > max ? t.slice(0, max - 1) + "…" : t;
};

/** Chữ ghi vào một ô của sheet nhật ký. Luôn là CHỮ: nội dung bắt đầu bằng = + - @ (Excel có thể hiểu là công thức) được thêm dấu nháy
 *  đầu để Excel không bao giờ chạy nó — nhật ký chép lại công thức và lời của model, không được tự tính hay gọi ra ngoài. */
export function logText(v: any, max = 250): string {
  if (v === null || v === undefined || v === "") return "";
  const t = oneLine(typeof v === "object" ? JSON.stringify(v) : String(v), max);
  return /^[=+\-@\t\r]/.test(t) ? "'" + t : t;
}

/** So hai khối cùng kích thước (trước / sau), trả các ô đổi nội dung. topLeft: cột, hàng tính từ 0. */
export function diffCells(before: any[][] | null | undefined, after: any[][], topLeft: { col: number; row: number }, addr: (col: number, row: number) => string): CellChange[] {
  const out: CellChange[] = [];
  const same = (a: any, b: any) => {
    if (a === b) return true;
    const na = a === null || a === undefined ? "" : a;
    const nb = b === null || b === undefined ? "" : b;
    if (na === nb) return true;
    // Số ghi dưới dạng chữ ("100") và số 100 là một
    return String(na) === String(nb);
  };
  after.forEach((row, i) =>
    row.forEach((v, j) => {
      const old = before?.[i]?.[j];
      if (!before || !same(old, v)) out.push({ cell: addr(topLeft.col + j, topLeft.row + i), before: before ? (old ?? "") : "(không đọc được)", after: v });
    })
  );
  return out;
}

/** Các dòng nhật ký của một thao tác. Mỗi ô đổi một dòng (tối đa MAX_CELL_ROWS, phần còn lại gộp); thao tác không đổi nội dung ô: một dòng. */
export function buildLogRows(o: LogInput): string[][] {
  const time = formatLogTime(o.time);
  const reason = o.reason?.trim() ? `${oneLine(o.action, 160)} — ${oneLine(o.reason, 240)}` : oneLine(o.action, 400);
  const tail = [logText(reason, 420), logText(o.request || "", 300)];
  const sheet = logText(o.sheet, 40);
  const changes = o.changes;
  if (!changes) return [[time, sheet, logText(o.address, 80), "", "", ...tail]];
  if (!changes.length) return [[time, sheet, logText(o.address, 80), "(không đổi)", "(không đổi)", ...tail]];
  const rows = changes.slice(0, MAX_CELL_ROWS).map((c) => [time, sheet, logText(c.cell, 20), logText(c.before), logText(c.after), ...tail]);
  if (changes.length > MAX_CELL_ROWS) {
    rows.push([time, sheet, logText(o.address, 80), "", `… và ${changes.length - MAX_CELL_ROWS} ô khác trong vùng này`, ...tail]);
  }
  return rows;
}

/** Dòng ghi khi người dùng bấm Hoàn tác một thao tác đã vào nhật ký. */
export function undoLogRow(time: Date, sheet: string, address: string, label: string): string[] {
  return [formatLogTime(time), logText(sheet, 40), logText(address, 80), "", "", logText(`Hoàn tác: ${label}`, 300), ""];
}

/** Tên sheet (đã bỏ dấu nháy bao ngoài) có phải sheet nhật ký không — so không phân biệt hoa thường, bỏ khoảng trắng hai đầu. */
export function isLogSheet(name: any, logSheet = LOG_SHEET): boolean {
  // Chuẩn hóa như công cụ làm sạch tên sheet trước khi dùng (ký tự \ / ? * [ ] : thành dấu cách, bỏ nháy bao ngoài, cắt 31 ký tự):
  // "Nhật ký AI?" hay "Nhật ký AI/" sau khi làm sạch chính là sheet nhật ký
  const norm = (x: any) =>
    String(x ?? "")
      .trim()
      .replace(/''/g, "'")
      .replace(/[\\/?*[\]:]/g, " ")
      .replace(/^'+|'+$/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 31)
      .trim()
      .toLowerCase();
  const n = norm(name);
  return !!n && n === norm(logSheet);
}

/** Như splitAreas của office/util (lõi không phụ thuộc lớp office): tách danh sách vùng, bỏ qua dấu phẩy trong tên sheet đặt trong nháy */
function areas(v: string): string[] {
  const out: string[] = [];
  let cur = "";
  let q = false;
  for (const ch of v) {
    if (ch === "'") q = !q;
    if (ch === "," && !q) {
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** Lệnh có nêu sheet ở tham số nào không (không nêu = chạy trên sheet đang mở: lớp bọc phải tự xem sheet đang mở có phải nhật ký không) */
export function namesSheet(args: any): boolean {
  if (!args || typeof args !== "object") return false;
  if (typeof args.sheet === "string" && args.sheet.trim()) return true;
  for (const k of ["address", "cell", "source_address", "fill_to", "print_area"]) {
    const v = args[k];
    if (typeof v === "string" && areas(v).some((p) => /^\s*(?:'(?:[^']|'')+'|[^'!]+)!/.test(p))) return true;
  }
  return false;
}

/** Lệnh công cụ có nhắm vào sheet nhật ký không (qua tham số sheet hoặc phần sheet của các tham số địa chỉ). */
export function targetsLogSheet(args: any, logSheet = LOG_SHEET): boolean {
  if (!args || typeof args !== "object") return false;
  if (isLogSheet(args.sheet, logSheet) || isLogSheet(args.new_name, logSheet) || isLogSheet(args.name, logSheet) || isLogSheet(args.sheet_name, logSheet)) return true;
  for (const k of ["address", "cell", "source_address", "destination", "fill_to", "print_area", "target_cell", "input_cell", "input_cell_2", "output_cell"]) {
    const v = args[k];
    if (typeof v !== "string") continue;
    for (const part of areas(v)) {
      const m = part.trim().match(/^(?:'((?:[^']|'')+)'|([^'!]+))!/);
      if (m && isLogSheet(m[1] ?? m[2], logSheet)) return true;
    }
  }
  return false;
}
