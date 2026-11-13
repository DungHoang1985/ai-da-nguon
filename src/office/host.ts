/* global Office, Excel, Word */
/** Làm việc với tài liệu đang mở: đọc nội dung đang chọn, chèn câu trả lời. */
import { parseMarkdownTable, stripMarkdown, valuesToTsv } from "../core/table";

export type HostName = "Excel" | "Word" | "PowerPoint" | "Khác";

export function hostName(): HostName {
  try {
    const h = Office.context.host;
    if (h === Office.HostType.Excel) return "Excel";
    if (h === Office.HostType.Word) return "Word";
    if (h === Office.HostType.PowerPoint) return "PowerPoint";
  } catch {
    /* chạy ngoài Office */
  }
  return "Khác";
}

export const INSERT_LABEL: Record<HostName, string> = {
  Excel: "Ghi vào ô đang chọn",
  Word: "Chèn vào chỗ con trỏ",
  PowerPoint: "Chèn vào slide",
  Khác: "Chèn",
};

function getSelectedTextAsync(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? String(r.value || "") : "");
    });
  });
}

function setSelectedTextAsync(text: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(r.error?.message || "Không chèn được"));
    });
  });
}

/** Đọc nội dung đang chọn để gửi kèm câu hỏi. */
export async function readSelection(): Promise<{ text: string; note: string }> {
  const host = hostName();
  if (host === "Excel") {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["address", "values"]);
      await ctx.sync();
      const { text, truncated } = valuesToTsv(range.values);
      return {
        text: `Vùng ${range.address}:\n${text}`,
        note: truncated ? "Vùng chọn lớn, đã cắt bớt còn 5.000 ô." : `Đã lấy vùng ${range.address}.`,
      };
    });
  }
  if (host === "Word") {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();
      let text = sel.text;
      let note = "Đã lấy đoạn đang chọn.";
      if (!text.trim()) {
        const body = ctx.document.body;
        body.load("text");
        await ctx.sync();
        text = body.text;
        note = "Không có vùng chọn — đã lấy toàn văn bản.";
      }
      const MAX = 60000;
      if (text.length > MAX) {
        text = text.slice(0, MAX);
        note += " Văn bản dài, đã cắt còn 60.000 ký tự.";
      }
      return { text, note };
    });
  }
  const text = await getSelectedTextAsync();
  return { text, note: text ? "Đã lấy chữ đang chọn." : "Không có chữ nào đang được chọn." };
}

/** Chèn câu trả lời vào tài liệu. Trả về mô tả ngắn đã làm gì. */
export async function insertAnswer(answer: string): Promise<string> {
  const host = hostName();
  if (host === "Excel") {
    const table = parseMarkdownTable(answer);
    return Excel.run(async (ctx) => {
      const cell = ctx.workbook.getSelectedRange().getCell(0, 0);
      if (table) {
        const target = cell.getResizedRange(table.length - 1, table[0].length - 1);
        target.values = table;
        target.load("address");
        await ctx.sync();
        return `Đã ghi bảng ${table.length}×${table[0].length} vào ${target.address}.`;
      }
      const lines = stripMarkdown(answer).split(/\r?\n/).filter((l) => l.trim() !== "");
      const target = cell.getResizedRange(Math.max(lines.length - 1, 0), 0);
      target.values = lines.map((l) => [l]);
      target.load("address");
      await ctx.sync();
      return `Đã ghi ${lines.length} dòng vào ${target.address}.`;
    });
  }
  if (host === "Word") {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.insertText(stripMarkdown(answer), Word.InsertLocation.after);
      await ctx.sync();
      return "Đã chèn vào văn bản.";
    });
  }
  await setSelectedTextAsync(stripMarkdown(answer));
  return "Đã chèn vào slide.";
}
