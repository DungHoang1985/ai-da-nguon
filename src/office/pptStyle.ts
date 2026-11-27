/**
 * Mức 3 việc 10 (27/11/2026, theo bản PowerPoint chính thức — `tasks/viec-le/rea-claude-chinh-thuc/` mục 6):
 * soát bằng mã chữ AI vừa ghi lên slide — văn phong lộ dấu AI, lời trao đổi lọt vào slide, cỡ chữ quá nhỏ, tiêu đề —
 * và trả cảnh báo trong KẾT QUẢ công cụ (không trong lời nhắc hệ thống, giữ trần ký tự). Chỉ lỗi khách quan bắt được bằng mẫu;
 * không chấm thẩm mỹ. Hàm thuần, không gọi Office.js.
 */

export const MIN_BODY_PT = 14;
/** Chú thích nguồn được phép nhỏ hơn (10–12 pt) */
export const MIN_NOTE_PT = 10;
export const MAX_TITLE_WORDS = 15;

export interface StyleWarning {
  /** Loại lỗi, ổn định để kiểm thử */
  kind: "emoji" | "contrast_frame" | "meta" | "dashes" | "hype" | "small_font" | "title_long" | "title_period" | "title_dup" | "no_key_message";
  /** Lời cho AI biết sửa gì */
  message: string;
  /** Đoạn chữ dính lỗi (≤ 80 ký tự) */
  excerpt?: string;
}

const cut = (s: string, n = 80) => (s.length > n ? s.slice(0, n - 1) + "…" : s);

/** Biểu tượng cảm xúc; trừ © ® ™ ‼ (tên thương hiệu, dự án trên slide pháp lý dùng ® ™ — phản biện mức 3) */
const EMOJI = /(?![©®™‼⁉])\p{Extended_Pictographic}/u;
/** Khuôn "Đây không phải X, mà là Y" / "không chỉ X mà còn Y" — dấu AI quen thuộc */
const CONTRAST = /(?<!\p{L})không (phải|chỉ)(?!\p{L})[^.\n]{1,80}?(?<!\p{L})mà (là|còn)(?!\p{L})/iu;
/** Lời trao đổi với người dùng lọt vào slide */
const META = /(?<!\p{L})([Tt]heo (góp ý|yêu cầu|ý) (của )?(bạn|anh|chị|người dùng)|[Bb]ản (rút gọn|chỉnh sửa|cập nhật) theo|[Nn]hư (bạn|anh|chị) (yêu cầu|muốn|đã nói)|[Tt]ôi đã (thêm|sửa|rút gọn|tạo))(?!\p{L})/u;
/** "AI đã / AI vừa" chỉ khi viết hoa (không bắt đại từ "ai": "ai đã ký") */
const META_AI = /(?<!\p{L})AI (đã|vừa)(?!\p{L})/u;
/** Từ nhấn rỗng (giữ ít, tránh bắt nhầm) */
const HYPE = /(?<!\p{L})(đột phá|vượt bậc|tuyệt vời|hàng đầu thế giới|không thể bỏ qua|cách mạng hóa)(?!\p{L})/iu;

/** Soát một khối chữ AI ghi lên slide */
export function lintSlideText(text: string): StyleWarning[] {
  const t = String(text || "");
  if (!t.trim()) return [];
  const out: StyleWarning[] = [];
  const e = t.match(EMOJI);
  if (e) out.push({ kind: "emoji", message: "Bỏ biểu tượng cảm xúc khỏi slide.", excerpt: cut(e[0]) });
  const c = t.match(CONTRAST);
  if (c) out.push({ kind: "contrast_frame", message: "Viết thẳng ý chính, bỏ khuôn \"không phải…, mà là…\" / \"không chỉ… mà còn…\".", excerpt: cut(c[0]) });
  const m = t.match(META) || t.match(META_AI);
  if (m) out.push({ kind: "meta", message: "Slide chứa lời trao đổi với người dùng; xóa câu này khỏi slide, chỉ nói trong chat.", excerpt: cut(m[0]) });
  // Chỉ đếm gạch ngang dùng như dấu câu (có khoảng trắng hai bên); "2020–2025", "Q1–Q3" là khoảng, không tính
  const dashes = (t.match(/\s[—–]\s/g) || []).length;
  if (dashes >= 3) out.push({ kind: "dashes", message: `Có ${dashes} gạch ngang dài; dùng dấu phẩy, hai chấm hoặc tách câu.` });
  const h = t.match(HYPE);
  if (h) out.push({ kind: "hype", message: "Bỏ từ nhấn rỗng, thay bằng số liệu hoặc sự kiện cụ thể.", excerpt: cut(h[0]) });
  return out;
}

/** Cỡ chữ AI đặt: thân ≥ 14 pt; chú thích nguồn ("Nguồn: …") được tới 10 pt */
export function lintFontSize(size: unknown, text?: string): StyleWarning[] {
  const n = Number(size);
  if (!Number.isFinite(n) || n <= 0) return [];
  const isNote = /^\s*(nguồn|source|ghi chú|chú thích)\s*[:：]/i.test(String(text || ""));
  const min = isNote ? MIN_NOTE_PT : MIN_BODY_PT;
  if (n >= min) return [];
  return [
    {
      kind: "small_font",
      message: isNote
        ? `Chú thích nguồn ${n} pt nhỏ hơn ${MIN_NOTE_PT} pt; đặt ${MIN_NOTE_PT}–12 pt.`
        : `Chữ ${n} pt nhỏ hơn ${MIN_BODY_PT} pt; không thu chữ để nhét — tách slide hoặc bớt ý.`,
    },
  ];
}

const words = (s: string) => s.trim().split(/\s+/).filter(Boolean).length;

/** Tiêu đề các slide của một lô: dài quá, chấm cuối, trùng nhau */
export function lintTitles(titles: string[]): StyleWarning[] {
  const out: StyleWarning[] = [];
  const seen = new Map<string, number>();
  titles.forEach((raw, i) => {
    const t = String(raw || "").trim();
    if (!t) return;
    if (words(t) > MAX_TITLE_WORDS) out.push({ kind: "title_long", message: `Tiêu đề slide thứ ${i + 1} dài ${words(t)} từ; rút còn ≤ ${MAX_TITLE_WORDS} từ.`, excerpt: cut(t) });
    if (/[.。]$/.test(t) && !/\.\.\.$/.test(t)) out.push({ kind: "title_period", message: `Tiêu đề slide thứ ${i + 1} có dấu chấm cuối; bỏ dấu chấm.`, excerpt: cut(t) });
    const k = t.toLowerCase();
    if (seen.has(k)) out.push({ kind: "title_dup", message: `Tiêu đề slide thứ ${i + 1} trùng slide thứ ${seen.get(k)! + 1}; mỗi slide một thông điệp riêng.`, excerpt: cut(t) });
    else seen.set(k, i);
  });
  return out;
}

/** Số slide của một lô từ đó cần câu ý chính (storyboard) */
export const STORYBOARD_MIN_SLIDES = 3;

/** Cảnh báo văn phong cho một lệnh ghi PowerPoint theo tham số AI gửi; [] nếu không có gì */
export function styleWarningsFor(name: string, a: any): StyleWarning[] {
  a = a || {};
  const out: StyleWarning[] = [];
  switch (name) {
    case "ppt_add_slides": {
      const list: any[] = Array.isArray(a.slides) ? a.slides : [];
      for (const s of list) out.push(...lintSlideText(`${s?.title || ""}\n${s?.body || ""}`));
      out.push(...lintTitles(list.map((s) => s?.title || "")));
      if (list.length >= STORYBOARD_MIN_SLIDES && !String(a.key_message || "").trim())
        out.push({ kind: "no_key_message", message: "Bài từ 3 slide: ghi key_message (một câu ý chính) để tiêu đề các slide cùng kể một câu chuyện." });
      break;
    }
    case "ppt_add_slide":
      out.push(...lintSlideText(`${a.title || ""}\n${a.body || ""}`), ...lintTitles([a.title || ""]));
      break;
    case "ppt_add_textbox":
    case "ppt_set_shape_text":
      out.push(...lintSlideText(a.text), ...lintFontSize(a.font_size, a.text));
      break;
    case "ppt_format_text":
      out.push(...lintFontSize(a.font_size, a.find));
      break;
    case "ppt_add_chart":
      out.push(...lintSlideText(`${a.title || ""}\n${a.chart_title || ""}`), ...lintTitles([a.title || ""]));
      break;
  }
  // (ranh giới từ viết bằng \p{L} và cờ u: \b của JavaScript chỉ hiểu chữ ASCII, bỏ sót từ bắt đầu bằng "đ")
  // Mỗi loại tối đa 3 lần, tổng tối đa 8 dòng (kết quả công cụ gọn)
  const per = new Map<string, number>();
  return out.filter((w) => (per.set(w.kind, (per.get(w.kind) || 0) + 1).get(w.kind) || 0) <= 3).slice(0, 8);
}

/** Gắn cảnh báo vào kết quả công cụ (đối tượng) — không đổi kết quả nếu không có cảnh báo */
export function withStyleWarnings<T>(name: string, a: any, result: T): T {
  const ws = styleWarningsFor(name, a);
  // Kết quả lỗi (ok:false), chuỗi, mảng: không gắn (phản biện mức 3)
  if (!ws.length || !result || typeof result !== "object" || Array.isArray(result) || (result as any).ok === false) return result;
  return { ...(result as any), style_warnings: ws.map((w) => (w.excerpt ? `${w.message} («${w.excerpt}»)` : w.message)) };
}
