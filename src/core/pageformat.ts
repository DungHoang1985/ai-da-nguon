/** Kiểm thể thức trang của một văn bản (lề, khổ giấy, phông và cỡ chữ) so với một chuẩn trình bày.
 *  Hàm thuần, không cần Office — kiểm thử bằng npm test. Số đo trang của Word tính bằng điểm (1 điểm = 1/72 inch). */

export const PT_PER_MM = 72 / 25.4;
export const ptToMm = (pt: number) => Math.round((pt / PT_PER_MM) * 10) / 10;

/** Thể thức trang của một phần (section) của văn bản, đơn vị điểm như Word trả về */
export interface PageSetupInfo {
  section: number;
  topMargin: number;
  bottomMargin: number;
  leftMargin: number;
  rightMargin: number;
  pageWidth: number;
  pageHeight: number;
  paperSize?: string;
  orientation?: string;
}

/** Phông và cỡ chữ của một đoạn. name rỗng hoặc size null = đoạn lẫn nhiều phông / nhiều cỡ */
export interface ParaFont {
  i: number;
  name: string;
  size: number | null;
  chars: number;
}

export interface PageStandard {
  name: string;
  /** Khổ giấy đứng, mm */
  paper: { w: number; h: number; label: string };
  /** Khoảng cho phép của từng lề, mm */
  top: [number, number];
  bottom: [number, number];
  left: [number, number];
  right: [number, number];
  /** Phông chữ được dùng (so không phân biệt hoa thường) */
  fonts: string[];
  /** Cỡ chữ của phần nội dung chính (cỡ chiếm nhiều ký tự nhất), điểm */
  bodySize: [number, number];
  /** Cỡ chữ nhỏ nhất và lớn nhất được phép ở bất kỳ đâu trong văn bản */
  anySize: [number, number];
}

/** Thể thức văn bản hành chính: giấy A4 đứng, lề trên và dưới 20–25 mm, lề trái 30–35 mm, lề phải 15–20 mm,
 *  phông Times New Roman, nội dung cỡ 13–14 (các thành phần khác từ cỡ 11 tới 14). */
export const VAN_BAN_HANH_CHINH: PageStandard = {
  name: "Thể thức văn bản hành chính",
  paper: { w: 210, h: 297, label: "A4" },
  top: [20, 25],
  bottom: [20, 25],
  left: [30, 35],
  right: [15, 20],
  fonts: ["Times New Roman"],
  bodySize: [13, 14],
  anySize: [11, 14],
};

export interface FormatIssue {
  /** paper | orientation | top | bottom | left | right | font | body_size | size | mixed */
  what: string;
  /** Tên dễ hiểu: "Lề trái" */
  label: string;
  actual: string;
  expected: string;
  section?: number;
  /** Đoạn đầu tiên có điểm này (phông, cỡ chữ) */
  p?: number;
  /** Các đoạn khác */
  also?: number[];
}

export interface FormatReport {
  standard: string;
  ok: boolean;
  issues: FormatIssue[];
  /** Số đo đọc được, đổi ra mm, để người đọc đối chiếu */
  sections: { section: number; paper: string; orientation: string; top_mm: number; bottom_mm: number; left_mm: number; right_mm: number }[];
  fonts: { name: string; size: number | null; chars: number; share: string }[];
}

const TOL_MM = 0.6;
const range = (r: [number, number], unit: string) => (r[0] === r[1] ? `${r[0]} ${unit}` : `${r[0]}–${r[1]} ${unit}`);
const mm = (v: number) => `${String(v).replace(".", ",")} mm`;

/** Chuẩn tùy chỉnh: ghi đè một phần của chuẩn gốc. Số lề là mm; một số = đúng giá trị đó, hai số = khoảng cho phép. */
export function customStandard(base: PageStandard, o: any): PageStandard {
  const pair = (v: any, def: [number, number]): [number, number] => {
    if (Array.isArray(v) && v.length >= 2 && v.every((x) => Number.isFinite(Number(x)))) return [Math.min(Number(v[0]), Number(v[1])), Math.max(Number(v[0]), Number(v[1]))];
    const n = Number(Array.isArray(v) ? v[0] : v);
    return v !== undefined && v !== null && v !== "" && Number.isFinite(n) && n >= 0 ? [n, n] : def;
  };
  if (!o || typeof o !== "object") return base;
  const fonts = Array.isArray(o.fonts) ? o.fonts.map(String).filter(Boolean) : typeof o.font === "string" && o.font.trim() ? [o.font.trim()] : base.fonts;
  const body = pair(o.font_size, base.bodySize);
  const custom = o.font_size !== undefined && o.font_size !== null && o.font_size !== "";
  return {
    ...base,
    name: "Chuẩn do người dùng nêu",
    top: pair(o.top_mm, base.top),
    bottom: pair(o.bottom_mm, base.bottom),
    left: pair(o.left_mm, base.left),
    right: pair(o.right_mm, base.right),
    fonts,
    bodySize: body,
    anySize: custom ? [Math.min(base.anySize[0], body[0]), Math.max(base.anySize[1], body[1])] : base.anySize,
  };
}

export function checkPageFormat(sections: PageSetupInfo[], paras: ParaFont[], std: PageStandard = VAN_BAN_HANH_CHINH): FormatReport {
  const issues: FormatIssue[] = [];
  const out: FormatReport["sections"] = [];
  const multi = sections.length > 1;
  for (const s of sections) {
    const w = ptToMm(s.pageWidth);
    const h = ptToMm(s.pageHeight);
    const landscape = w > h + TOL_MM;
    const short = Math.min(w, h);
    const long = Math.max(w, h);
    const paperOk = Math.abs(short - std.paper.w) <= 1 && Math.abs(long - std.paper.h) <= 1;
    const paper = paperOk ? std.paper.label : s.paperSize && s.paperSize !== "Custom" ? `${s.paperSize} (${mm(w)} × ${mm(h)})` : `${mm(w)} × ${mm(h)}`;
    const row = { section: s.section, paper, orientation: landscape ? "ngang" : "đứng", top_mm: ptToMm(s.topMargin), bottom_mm: ptToMm(s.bottomMargin), left_mm: ptToMm(s.leftMargin), right_mm: ptToMm(s.rightMargin) };
    out.push(row);
    const sec = multi ? { section: s.section } : {};
    if (!paperOk) issues.push({ what: "paper", label: "Khổ giấy", actual: paper, expected: `${std.paper.label} (${std.paper.w} × ${std.paper.h} mm)`, ...sec });
    else if (landscape) issues.push({ what: "orientation", label: "Hướng giấy", actual: "ngang", expected: "đứng", ...sec });
    const margin = (what: "top" | "bottom" | "left" | "right", label: string, v: number) => {
      const r = std[what];
      if (v < r[0] - TOL_MM || v > r[1] + TOL_MM) issues.push({ what, label, actual: mm(v), expected: range(r, "mm"), ...sec });
    };
    margin("top", "Lề trên", row.top_mm);
    margin("bottom", "Lề dưới", row.bottom_mm);
    margin("left", "Lề trái", row.left_mm);
    margin("right", "Lề phải", row.right_mm);
  }

  // Phông và cỡ chữ: gom theo số ký tự
  const text = paras.filter((p) => p.chars > 0);
  const total = text.reduce((n, p) => n + p.chars, 0) || 1;
  const byFont = new Map<string, { name: string; size: number | null; chars: number; ps: number[] }>();
  for (const p of text) {
    const k = `${p.name}|${p.size}`;
    const g = byFont.get(k) || { name: p.name, size: p.size, chars: 0, ps: [] };
    g.chars += p.chars;
    g.ps.push(p.i);
    byFont.set(k, g);
  }
  const groups = [...byFont.values()].sort((a, b) => b.chars - a.chars);
  const okFont = (name: string) => std.fonts.some((f) => f.toLowerCase() === name.toLowerCase());
  const seenFont = new Set<string>();
  for (const g of groups) {
    if (g.name && !okFont(g.name) && !seenFont.has(g.name)) {
      seenFont.add(g.name);
      const ps = groups.filter((x) => x.name === g.name).flatMap((x) => x.ps).sort((a, b) => a - b);
      issues.push({ what: "font", label: "Phông chữ", actual: `${g.name} (${ps.length} đoạn)`, expected: std.fonts.join(" hoặc "), p: ps[0], ...(ps.length > 1 ? { also: ps.slice(1, 12) } : {}) });
    }
  }
  const mixed = text.filter((p) => !p.name || p.size === null).map((p) => p.i);
  if (mixed.length) {
    issues.push({ what: "mixed", label: "Đoạn lẫn nhiều phông hoặc nhiều cỡ chữ", actual: `${mixed.length} đoạn`, expected: "một phông, một cỡ trong mỗi đoạn", p: mixed[0], ...(mixed.length > 1 ? { also: mixed.slice(1, 12) } : {}) });
  }
  const sized = groups.filter((g) => g.size !== null);
  if (sized.length) {
    // Cỡ chữ của nội dung chính = cỡ chiếm nhiều ký tự nhất
    const bySize = new Map<number, number>();
    sized.forEach((g) => bySize.set(g.size as number, (bySize.get(g.size as number) || 0) + g.chars));
    const main = [...bySize.entries()].sort((a, b) => b[1] - a[1])[0][0];
    if (main < std.bodySize[0] || main > std.bodySize[1]) {
      const ps = sized.filter((g) => g.size === main).flatMap((g) => g.ps).sort((a, b) => a - b);
      issues.push({ what: "body_size", label: "Cỡ chữ nội dung", actual: `${main}`, expected: range(std.bodySize, "").trim(), p: ps[0] });
    }
    for (const [size] of [...bySize.entries()].sort((a, b) => a[0] - b[0])) {
      if (size === main || (size >= std.anySize[0] && size <= std.anySize[1])) continue;
      const ps = sized.filter((g) => g.size === size).flatMap((g) => g.ps).sort((a, b) => a - b);
      issues.push({ what: "size", label: "Cỡ chữ", actual: `${size} (${ps.length} đoạn)`, expected: range(std.anySize, "").trim(), p: ps[0], ...(ps.length > 1 ? { also: ps.slice(1, 12) } : {}) });
    }
  }
  return {
    standard: std.name,
    ok: issues.length === 0,
    issues,
    sections: out,
    fonts: groups.slice(0, 8).map((g) => ({ name: g.name || "(lẫn nhiều phông)", size: g.size, chars: g.chars, share: `${Math.round((g.chars / total) * 100)}%` })),
  };
}
