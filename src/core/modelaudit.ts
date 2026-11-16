/** Kiểm toán mô hình bảng tính bằng mã (Đợt 4), không tốn token: hằng số gõ cứng nằm giữa dãy công thức, công thức lệch so với
 *  hai bên, tỷ lệ gõ thẳng trong công thức, liên kết ra tệp ngoài. Hàm thuần: nhận lưới công thức kiểu R1C1 (cùng một công thức kéo
 *  sang ô khác thì chuỗi R1C1 giống hệt nhau) và trả "điểm cần xem". Ưu tiên không báo nhầm: chỉ xét ô nằm GIỮA dãy (đầu và cuối
 *  dãy thường là nhãn, số đầu vào hoặc dòng tổng). */
import { formulaRisksOf, stripStrings } from "./formularisk";

export type AuditKind = "hardcoded" | "inconsistent" | "embedded_rate" | "external_link";

export interface AuditFinding {
  kind: AuditKind;
  /** Ô đầu tiên của phát hiện, ví dụ "D6" */
  cell: string;
  /** Số ô cùng phát hiện (cùng một mẫu công thức) */
  count: number;
  /** Các ô (tối đa 12) */
  cells: string[];
  detail: string;
}

export interface AuditInput {
  /** formulasR1C1 của vùng: chuỗi bắt đầu bằng "=" là công thức, số là hằng số, còn lại là chữ / trống */
  r1c1: any[][];
  /** formulas kiểu A1 (để hiện cho người đọc); thiếu thì dùng r1c1 */
  formulas?: any[][];
  /** Ô trên-trái của vùng: cột, hàng tính từ 0 */
  origin: { col: number; row: number };
  addr: (col: number, row: number) => string;
}

export interface AuditResult {
  findings: AuditFinding[];
  counts: Record<AuditKind, number>;
  scanned: { cells: number; formulas: number; constants: number };
}

const isF = (v: any): v is string => typeof v === "string" && v.startsWith("=");
const isNum = (v: any) => typeof v === "number" && Number.isFinite(v);

/** Số gõ thẳng trong công thức đáng xem: số thập phân (0.05, 1.1) hoặc phần trăm (10%). Số nguyên (1, 12, 100, 365) bỏ qua vì
 *  thường là hằng số kỹ thuật (số tháng, làm tròn, chỉ số cột). Chuỗi và tên sheet trong nháy không tính. */
export function embeddedRates(formula: string): string[] {
  // Bỏ chuỗi, tên sheet trong nháy và phần trong ngoặc vuông (tên cột của bảng như [VAT 10%], [Phí bảo trì 2%])
  let code = stripStrings(String(formula ?? "")).replace(/'(?:[^']|'')*'(?=\s*!)/g, "S");
  for (let k = 0; k < 4 && /\[[^\[\]]*\]/.test(code); k++) code = code.replace(/\[[^\[\]]*\]/g, "T");
  const out: string[] = [];
  // Dấu phẩy là dấu ngăn tham số, không phải dấu thập phân: chỉ nhận số có dấu chấm, hoặc số kèm %
  const re = /(?<![\p{L}_$.\d])(\d+\.\d+%?|\d+%)(?![\d.]*[\p{L}_(!:])/gu;
  let m: RegExpExecArray | null;
  while ((m = re.exec(code))) out.push(m[1]);
  return [...new Set(out)];
}

export function auditModel(input: AuditInput): AuditResult {
  const g = input.r1c1 || [];
  const a1 = input.formulas || g;
  const rows = g.length;
  const cols = rows ? Math.max(...g.map((r) => (Array.isArray(r) ? r.length : 0))) : 0;
  const at = (i: number, j: number) => (i >= 0 && j >= 0 && i < rows && j < cols ? g[i]?.[j] : undefined);
  const name = (i: number, j: number) => input.addr(input.origin.col + j, input.origin.row + i);
  const show = (i: number, j: number) => {
    const f = String(a1[i]?.[j] ?? g[i]?.[j] ?? "");
    return f.length > 80 ? f.slice(0, 79) + "…" : f;
  };
  const findings: AuditFinding[] = [];
  let formulas = 0;
  let constants = 0;
  const embedded = new Map<string, { cells: [number, number][]; rates: string[] }>();
  const external: [number, number][] = [];

  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const v = at(i, j);
      const L = at(i, j - 1);
      const R = at(i, j + 1);
      const U = at(i - 1, j);
      const D = at(i + 1, j);
      if (isNum(v)) {
        constants++;
        // Hằng số nằm giữa hai công thức cùng mẫu (theo hàng hoặc theo cột): nhiều khả năng một ô công thức bị gõ đè.
        // Nhưng nếu ô liền kề theo chiều kia cũng là số nhập tay thì đây là cả một hàng / cột số liệu nhập (hàng "Số căn" giữa hai hàng
        // công thức, cột "Thực hiện" xen giữa các cột "Kế hoạch", hàng tiêu đề năm): không phải gõ đè — soát mã 16/11/2026, #13
        const h = isF(L) && isF(R) && L === R && !isNum(U) && !isNum(D);
        const vv = isF(U) && isF(D) && U === D && !isNum(L) && !isNum(R);
        if (h || vv) {
          const [pi, pj] = h ? [i, j - 1] : [i - 1, j];
          findings.push({
            kind: "hardcoded",
            cell: name(i, j),
            count: 1,
            cells: [name(i, j)],
            detail: `Ô ${name(i, j)} là số gõ cứng (${v}) nằm giữa dãy công thức theo ${h ? "hàng" : "cột"}; hai ô bên cạnh cùng công thức ${show(pi, pj)}`,
          });
        }
        continue;
      }
      if (!isF(v)) continue;
      formulas++;
      // Lệch so với hai bên: hai bên cùng mẫu, mẫu đó còn lặp thêm ít nhất một ô nữa (dãy thật), và ô này không khớp với dãy theo chiều kia
      const hRun = isF(L) && isF(R) && L === R && (at(i, j - 2) === L || at(i, j + 2) === R);
      const vRun = isF(U) && isF(D) && U === D && (at(i - 2, j) === U || at(i + 2, j) === D);
      const hOdd = hRun && v !== L;
      const vOdd = vRun && v !== U;
      const hMatch = v === L || v === R;
      const vMatch = v === U || v === D;
      if ((hOdd && !vMatch) || (vOdd && !hMatch)) {
        const [pi, pj] = hOdd && !vMatch ? [i, j - 1] : [i - 1, j];
        findings.push({
          kind: "inconsistent",
          cell: name(i, j),
          count: 1,
          cells: [name(i, j)],
          detail: `Công thức ở ${name(i, j)} (${show(i, j)}) lệch so với các ô cùng ${hOdd && !vMatch ? "hàng" : "cột"} (ô bên cạnh: ${show(pi, pj)})`,
        });
      }
      const rates = embeddedRates(String(a1[i]?.[j] ?? v));
      if (rates.length) {
        const e = embedded.get(v) || { cells: [], rates };
        e.cells.push([i, j]);
        embedded.set(v, e);
      }
      if (formulaRisksOf(String(a1[i]?.[j] ?? v)).some((x) => x.what === "liên kết tệp ngoài")) external.push([i, j]);
    }
  }
  for (const e of embedded.values()) {
    const [i, j] = e.cells[0];
    findings.push({
      kind: "embedded_rate",
      cell: name(i, j),
      count: e.cells.length,
      cells: e.cells.slice(0, 12).map(([x, y]) => name(x, y)),
      detail: `Tỷ lệ ${e.rates.join(", ")} gõ thẳng trong công thức ${show(i, j)}${e.cells.length > 1 ? ` (${e.cells.length} ô cùng mẫu)` : ""}: nên đưa ra một ô giả định riêng`,
    });
  }
  if (external.length) {
    const [i, j] = external[0];
    findings.push({
      kind: "external_link",
      cell: name(i, j),
      count: external.length,
      cells: external.slice(0, 12).map(([x, y]) => name(x, y)),
      detail: `Công thức lấy số từ tệp khác, ví dụ ${name(i, j)}: ${show(i, j)}`,
    });
  }
  const counts: Record<AuditKind, number> = { hardcoded: 0, inconsistent: 0, embedded_rate: 0, external_link: 0 };
  for (const f of findings) counts[f.kind] += f.kind === "hardcoded" || f.kind === "inconsistent" ? 1 : f.count;
  return { findings, counts, scanned: { cells: rows * cols, formulas, constants } };
}
