/** Tách bảng Markdown đầu tiên trong câu trả lời thành mảng 2 chiều (để ghi vào Excel). */
export function parseMarkdownTable(text: string): string[][] | null {
  const lines = text.split(/\r?\n/);
  for (let i = 0; i + 1 < lines.length; i++) {
    const header = lines[i].trim();
    const sep = lines[i + 1].trim();
    if (header.includes("|") && /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$/.test(sep)) {
      const rows: string[][] = [splitRow(header)];
      for (let j = i + 2; j < lines.length; j++) {
        const l = lines[j].trim();
        if (!l.includes("|")) break;
        rows.push(splitRow(l));
      }
      const width = Math.max(...rows.map((r) => r.length));
      return rows.map((r) => [...r, ...Array(width - r.length).fill("")]);
    }
  }
  return null;
}

function splitRow(line: string): string[] {
  let l = line.trim();
  if (l.startsWith("|")) l = l.slice(1);
  if (l.endsWith("|")) l = l.slice(0, -1);
  return l.split("|").map((c) => c.trim().replace(/\*\*(.*?)\*\*/g, "$1"));
}

/** Bỏ ký hiệu Markdown đơn giản khi chèn chữ thường vào tài liệu. */
export function stripMarkdown(text: string): string {
  return text
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/__(.*?)__/g, "$1")
    .replace(/`{3}[a-zA-Z]*\n?/g, "")
    .replace(/`([^`]*)`/g, "$1")
    .replace(/^\s*[-*]\s+/gm, "• ")
    .trim();
}

/** Chuyển ô Excel đang chọn thành văn bản dạng bảng (tab) để gửi cho AI. */
export function valuesToTsv(values: any[][], maxCells = 5000): { text: string; truncated: boolean } {
  let count = 0;
  const out: string[] = [];
  let truncated = false;
  for (const row of values) {
    if (count + row.length > maxCells) {
      truncated = true;
      break;
    }
    count += row.length;
    out.push(row.map((v) => (v === null || v === undefined ? "" : String(v))).join("\t"));
  }
  return { text: out.join("\n"), truncated };
}
