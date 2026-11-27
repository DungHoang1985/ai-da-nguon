/** Lý do AI nêu cho lệnh sửa Word — tách khỏi wordTools.ts (B10 phần 2, 27/11/2026) để nhãn bước trên khung chat
 *  không kéo cả bộ công cụ Word vào gói tải đầu của Excel, PowerPoint. */

/** Công cụ sửa nội dung Word nhận thêm `reason`: một câu lý do ngắn, hiện cạnh bước trên khung, trong hộp hỏi và ở mục Hoàn tác
 *  (như thẻ đề xuất kèm lý do của add-in Claude chính thức). Lý do chỉ để hiển thị, không ghi vào tài liệu. */
export const REASON_TOOLS = new Set(["word_replace_text", "word_set_paragraph_text", "word_insert_paragraphs", "word_insert_clause", "word_replace_selection", "word_delete_paragraphs", "word_fill_table"]);

/** Lý do AI nêu cho một lần sửa, gọn một dòng (rỗng nếu không nêu) */
export function reasonOf(name: string, a: any): string {
  if (!REASON_TOOLS.has(name)) return "";
  const r = String(a?.reason ?? "").replace(/\s+/g, " ").trim();
  return r.length > 160 ? r.slice(0, 159) + "…" : r;
}
