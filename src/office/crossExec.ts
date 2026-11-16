/** Khung NHẬN của ghi chéo ứng dụng: chạy một thao tác ghi hộ (đã qua kiểm ở core/crossapp.ts) bằng chính công cụ ghi của ứng dụng,
 *  nên Word vẫn ghi bằng Theo dõi thay đổi và mọi thao tác đều có Hoàn tác trong nhật ký của khung này.
 *  Nội dung nhận được chỉ là dữ liệu để chèn: không đưa cho model, không chạy như chỉ dẫn. */
import { HostToolSet, UndoEntry } from "../core/agent/types";
import { CrossOp } from "../core/crossapp";

/** Ô bắt đầu bằng =, + hoặc @ được ghi thành chữ: không nhận công thức từ ứng dụng khác (dấu nháy đầu ô = kiểu chữ, theo quy ước của excel_write_range) */
export const cellAsText = (v: string): string => (/^[=+@]/.test(v) || (/^-/.test(v) && !/^-\d[\d.,\s]*%?$/.test(v)) ? "'" + v : v);

export async function executeIncoming(tools: HostToolSet, p: CrossOp, log: UndoEntry[]): Promise<string> {
  const run = (name: string, args: any) => tools.execute(name, args, { log });
  if (p.op === "word_append") {
    if (tools.host !== "Word") throw new Error("Thao tác này chỉ dành cho Word.");
    const r = await run("word_insert_paragraphs", { text: p.text, position: "end" });
    return r?.message || "Đã thêm nội dung vào cuối văn bản (ghi có theo dõi).";
  }
  if (p.op === "excel_new_sheet") {
    if (tools.host !== "Excel") throw new Error("Thao tác này chỉ dành cho Excel.");
    let name = p.name;
    let lastErr: any;
    for (let k = 1; k <= 4; k++) {
      try {
        const made = await run("excel_add_sheet", { name, activate: false });
        if (made?.sheet) name = made.sheet;
        lastErr = null;
        break;
      } catch (e) {
        lastErr = e;
        name = `${p.name.slice(0, 25)} (${k + 1})`;
      }
    }
    if (lastErr) throw lastErr;
    await run("excel_write_range", { sheet: name, address: "A1", values: p.values.map((r) => r.map(cellAsText)) });
    return `Đã tạo sheet “${name}” và ghi bảng ${p.values.length}×${p.values[0]?.length || 0} ô (không ghi đè dữ liệu có sẵn).`;
  }
  if (tools.host !== "PowerPoint") throw new Error("Thao tác này chỉ dành cho PowerPoint.");
  const r = await run("ppt_add_slide", { title: p.title, body: p.body });
  return r?.message || "Đã thêm một slide ở cuối bài.";
}
