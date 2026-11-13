/** Tìm địa chỉ ô / đoạn / slide trong câu trả lời để biến thành liên kết bấm được. */

export type RefTarget =
  | { kind: "cell"; address: string }
  | { kind: "para"; index: number }
  | { kind: "slide"; index: number };

export type RefPart = { text: string; ref?: RefTarget };

const CELL =
  /(?:'[^'\n]{1,60}'|[A-Za-z0-9_À-ỹ.]{1,40})!\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?|\b\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?\b/g;
const PARA = /\[¶\s?(\d{1,5})\]|¶\s?(\d{1,5})|(?<![A-Za-zÀ-ỹ])đoạn\s+(?:số\s+)?(\d{1,5})\b/gi;
const SLIDE = /\bslide\s+(?:số\s+)?(\d{1,3})\b/gi;

function split(text: string, re: RegExp, make: (m: RegExpExecArray) => RefTarget | null): RefPart[] {
  const out: RefPart[] = [];
  let last = 0;
  re.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const ref = make(m);
    if (!ref) continue;
    if (m.index > last) out.push({ text: text.slice(last, m.index) });
    out.push({ text: m[0], ref });
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push({ text: text.slice(last) });
  return out;
}

/** Ô kiểu "A1" phải trông như địa chỉ thật (cột ≤ XFD, hàng ≥ 1) và không nằm trong từ */
function plausibleCell(s: string): boolean {
  const a = s.split("!").pop()!.replace(/\$/g, "");
  return a.split(":").every((p) => {
    const m = p.match(/^([A-Z]{1,3})(\d{1,7})$/);
    if (!m) return false;
    const col = m[1];
    if (col.length === 3 && col > "XFD") return false;
    return Number(m[2]) >= 1 && Number(m[2]) <= 1048576;
  });
}

export function findRefs(text: string, host: string): RefPart[] {
  if (!text) return [];
  if (host === "Excel") {
    return split(text, CELL, (m) => (plausibleCell(m[0]) ? { kind: "cell", address: m[0].replace(/\$/g, "") } : null));
  }
  if (host === "Word") {
    return split(text, PARA, (m) => ({ kind: "para", index: Number(m[1] ?? m[2] ?? m[3]) }));
  }
  if (host === "PowerPoint") {
    return split(text, SLIDE, (m) => ({ kind: "slide", index: Number(m[1]) }));
  }
  return [{ text }];
}
