/** Kiểm bài trình chiếu theo mẫu bằng mã (Đợt 3, việc B3) — hàm thuần, không tốn token:
 *  phông lệch chuẩn, chữ quá nhỏ, slide thiếu tiêu đề, chữ ước lượng tràn khung, hình ra ngoài slide, khung chữ chồng nhau,
 *  chữ mẫu còn sót, màu chữ ngoài bảng màu chủ đề. Kết quả là "điểm cần xem"; riêng chữ tràn là ước lượng (không đo được chữ thật). */

export interface DeckPara {
  text: string;
  /** Tên phông; null = đoạn dùng lẫn nhiều phông */
  font: string | null;
  /** Cỡ chữ (pt); null = lẫn nhiều cỡ */
  size: number | null;
  color?: string | null;
}

export interface DeckShape {
  name: string;
  /** Số thứ tự của hình trên slide (1-based) — để chỉ đúng hình khi nhiều hình trùng tên */
  n?: number;
  type: string;
  /** Ô giữ chỗ tiêu đề của slide */
  isTitle?: boolean;
  /** left, top, width, height (pt) */
  box: [number, number, number, number];
  paragraphs?: DeckPara[];
  /** Khung tự giãn theo chữ (không tràn) hoặc chữ tự co */
  autoSize?: string;
  wordWrap?: boolean;
}

export interface DeckSlide {
  slide: number;
  layout?: string;
  shapes: DeckShape[];
}

export interface DeckOptions {
  slideW: number;
  slideH: number;
  /** Phông của mẫu (tiêu đề, nội dung). Rỗng = lấy phông dùng nhiều nhất trong bài làm chuẩn */
  themeFonts?: string[];
  /** Màu chủ đề (#RRGGBB) */
  themeColors?: string[];
  /** Cỡ chữ nhỏ nhất chấp nhận (mặc định 12) */
  minSize?: number;
}

export type DeckIssueKind = "missing_title" | "font" | "small_text" | "overflow" | "off_slide" | "overlap" | "placeholder_text" | "color";

export interface DeckIssue {
  kind: DeckIssueKind;
  slide: number;
  shape?: string;
  message: string;
  /** Gợi ý sửa bằng công cụ */
  fix?: string;
}

export interface DeckReport {
  standardFonts: string[];
  minSize: number;
  slides: number;
  issues: DeckIssue[];
  counts: Partial<Record<DeckIssueKind, number>>;
}

const LABEL: Record<DeckIssueKind, string> = {
  missing_title: "slide thiếu tiêu đề",
  font: "phông lệch mẫu",
  small_text: "chữ quá nhỏ",
  overflow: "chữ tràn khung",
  off_slide: "hình ra ngoài slide",
  overlap: "khung chữ chồng nhau",
  placeholder_text: "chữ mẫu còn sót",
  color: "màu chữ ngoài bảng màu chủ đề",
};
export const deckIssueLabel = (k: DeckIssueKind) => LABEL[k];

const norm = (f: string) => f.trim().toLowerCase();
const cut = (s: string, n = 40) => (s.length > n ? s.slice(0, n) + "…" : s);
const PLACEHOLDER_TEXT = /^(click to add|bấm để thêm|nhấp để thêm|lorem ipsum|\[cần bổ sung\]|xxx+|tbd)\b/i;

/** Phông dùng nhiều nhất (theo số ký tự) */
export function dominantFonts(slides: DeckSlide[], top = 1): string[] {
  const n = new Map<string, number>();
  for (const s of slides) for (const sh of s.shapes) for (const p of sh.paragraphs || []) if (p.font && p.text.trim()) n.set(p.font, (n.get(p.font) || 0) + p.text.length);
  return [...n.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, top)
    .map((x) => x[0]);
}

/** Chiều cao chữ ước lượng của một khung (pt): mỗi ký tự rộng khoảng nửa cỡ chữ, dòng cao 1,2 lần cỡ chữ */
export function estimateTextHeight(paras: DeckPara[], innerWidth: number, fallbackSize = 18): number {
  let h = 0;
  for (const p of paras) {
    const size = p.size || fallbackSize;
    const perLine = Math.max(1, Math.floor(innerWidth / (size * 0.5)));
    const lines = Math.max(1, Math.ceil(p.text.length / perLine));
    h += lines * size * 1.2;
  }
  return h;
}

export function checkDeck(slides: DeckSlide[], o: DeckOptions): DeckReport {
  const minSize = o.minSize && o.minSize > 0 ? o.minSize : 12;
  const standard = (o.themeFonts && o.themeFonts.filter(Boolean).length ? o.themeFonts.filter(Boolean) : dominantFonts(slides, 1)).map((f) => f.trim());
  const std = new Set(standard.map(norm));
  const palette = new Set((o.themeColors || []).map((c) => c.toUpperCase()));
  const neutral = /^#(000000|FFFFFF|[0-9A-F]{6})$/i;
  const issues: DeckIssue[] = [];
  const add = (i: DeckIssue) => issues.push(i);
  /** Cách gọi một hình trong câu báo: kèm số thứ tự để AI chỉ đúng hình khi tên trùng nhau */
  const q = (sh: DeckShape) => (sh.n ? `“${sh.name}” (hình số ${sh.n})` : `“${sh.name}”`);
  /** Vùng chữ thật của một khung (ước lượng, chữ neo trên): khung rộng mà ít chữ thì không tính phần trống là "chồng" */
  const textRect = (sh: DeckShape): [number, number, number, number] => {
    const [l, t, w, h] = sh.box;
    const paras = (sh.paragraphs || []).filter((p) => p.text.trim());
    const need = estimateTextHeight(paras, Math.max(10, w - 14.4)) + 7.2;
    return [l, t, w, Math.min(h, need)];
  };

  for (const s of slides) {
    const texty = s.shapes.filter((sh) => (sh.paragraphs || []).some((p) => p.text.trim()));
    // Tiêu đề
    const title = s.shapes.find((sh) => sh.isTitle);
    if (!title || !(title.paragraphs || []).some((p) => p.text.trim())) {
      add({ kind: "missing_title", slide: s.slide, message: `Slide ${s.slide} ${title ? "có ô tiêu đề nhưng để trống" : "không có tiêu đề"}.`, fix: title ? "ppt_set_shape_text vào ô tiêu đề" : "ppt_set_layout sang bố cục có tiêu đề rồi điền tiêu đề" });
    }
    for (const sh of texty) {
      const paras = (sh.paragraphs || []).filter((p) => p.text.trim());
      // Phông: gộp theo hình, mỗi phông lạ một điểm
      const odd = new Map<string, string>();
      for (const p of paras) {
        const key = p.font === null ? "(lẫn nhiều phông trong một đoạn)" : p.font;
        if (std.size && (p.font === null || !std.has(norm(p.font))) && !odd.has(key)) odd.set(key, p.text);
      }
      for (const [font, sample] of odd)
        add({ kind: "font", slide: s.slide, shape: sh.name, message: `Slide ${s.slide}, ${q(sh)}: phông ${font} (mẫu dùng ${standard.join(", ")}) — “${cut(sample.trim())}”.`, fix: `ppt_format_text với font_name "${standard[standard.length - 1]}"` });
      // Cỡ chữ: mỗi hình một điểm, lấy cỡ nhỏ nhất
      const small = paras.filter((p) => p.size !== null && p.size < minSize).sort((a, b) => (a.size as number) - (b.size as number));
      if (small.length)
        add({ kind: "small_text", slide: s.slide, shape: sh.name, message: `Slide ${s.slide}, ${q(sh)}: chữ cỡ ${small[0].size} pt (nhỏ hơn ${minSize} pt) — “${cut(small[0].text.trim())}”.`, fix: `ppt_format_text với font_size từ ${minSize} trở lên, hoặc tách bớt nội dung sang slide khác` });
      // Chữ mẫu còn sót
      const ph = paras.find((p) => PLACEHOLDER_TEXT.test(p.text.trim()));
      if (ph) add({ kind: "placeholder_text", slide: s.slide, shape: sh.name, message: `Slide ${s.slide}, ${q(sh)}: còn chữ mẫu “${cut(ph.text.trim())}”.` });
      // Màu chữ
      if (palette.size) {
        const bad = paras.find((p) => p.color && neutral.test(p.color) && !palette.has(p.color.toUpperCase()) && !/^#(000000|FFFFFF)$/i.test(p.color));
        if (bad) add({ kind: "color", slide: s.slide, shape: sh.name, message: `Slide ${s.slide}, ${q(sh)}: màu chữ ${bad.color} không thuộc bảng màu chủ đề.` });
      }
      // Chữ tràn khung (ước lượng): chỉ với khung không tự giãn / tự co
      const [, , w, h] = sh.box;
      const fixed = !sh.autoSize || sh.autoSize === "AutoSizeNone";
      if (fixed && sh.wordWrap !== false && w > 20 && h > 0) {
        const need = estimateTextHeight(paras, Math.max(10, w - 14.4)) + 7.2;
        if (need > h * 1.15 + 4) add({ kind: "overflow", slide: s.slide, shape: sh.name, message: `Slide ${s.slide}, ${q(sh)}: chữ ước cao ${Math.round(need)} pt trong khung cao ${Math.round(h)} pt — có thể tràn khung.`, fix: "giảm cỡ chữ (ppt_format_text), nới khung (ppt_update_shape) hoặc tách nội dung; rồi chụp lại slide để xem" });
      }
    }
    // Ra ngoài slide
    for (const sh of s.shapes) {
      const [l, t, w, h] = sh.box;
      if (w <= 0 || h <= 0) continue;
      if (l < -2 || t < -2 || l + w > o.slideW + 2 || t + h > o.slideH + 2)
        add({ kind: "off_slide", slide: s.slide, shape: sh.name, message: `Slide ${s.slide}, ${q(sh)}: nằm ngoài slide (trái ${Math.round(l)}, trên ${Math.round(t)}, rộng ${Math.round(w)}, cao ${Math.round(h)}; slide ${Math.round(o.slideW)}×${Math.round(o.slideH)}).`, fix: "ppt_update_shape đặt lại left, top, width, height" });
    }
    // Chữ của hai khung chồng nhau (đáng kể: phần chồng quá 25% vùng chữ nhỏ hơn). So theo vùng chữ ước lượng, không theo cả khung.
    for (let i = 0; i < texty.length; i++)
      for (let j = i + 1; j < texty.length; j++) {
        const [al, at, aw, ah] = textRect(texty[i]);
        const [bl, bt, bw, bh] = textRect(texty[j]);
        const ox = Math.min(al + aw, bl + bw) - Math.max(al, bl);
        const oy = Math.min(at + ah, bt + bh) - Math.max(at, bt);
        if (ox > 0 && oy > 0 && ox * oy > 0.25 * Math.min(aw * ah, bw * bh))
          add({ kind: "overlap", slide: s.slide, shape: texty[i].name, message: `Slide ${s.slide}: chữ của ${q(texty[i])} và ${q(texty[j])} chồng lên nhau.`, fix: "ppt_update_shape dời hoặc thu nhỏ một trong hai" });
      }
  }
  const counts: Partial<Record<DeckIssueKind, number>> = {};
  for (const i of issues) counts[i.kind] = (counts[i.kind] || 0) + 1;
  return { standardFonts: standard, minSize, slides: slides.length, issues, counts };
}

/** Một dòng tóm tắt hiện trên khung */
export function deckBrief(r: DeckReport): string {
  if (!r.issues.length) return `${r.slides} slide, không thấy điểm lệch mẫu`;
  return `${r.issues.length} điểm cần xem: ` + (Object.keys(r.counts) as DeckIssueKind[]).map((k) => `${r.counts[k]} ${LABEL[k]}`).join(", ");
}
