/** Kho tạm trong bộ nhớ cho bản gốc của các tệp .docx người dùng đính kèm trong phiên làm việc này, để công cụ Word so sánh
 *  phiên bản đọc được tệp. Không lưu xuống đĩa; tải lại khung add-in là mất (khi đó công cụ nhắc người dùng đính kèm lại). */

export interface AttachedFile {
  name: string;
  /** base64 của tệp gốc */
  data: string;
  at: number;
}

const MAX_FILES = 6;
const files: AttachedFile[] = [];

export function registerAttachedFile(name: string, data: string, now = Date.now()): void {
  const at = files.findIndex((f) => f.name === name);
  if (at >= 0) files.splice(at, 1);
  files.push({ name, data, at: now });
  while (files.length > MAX_FILES) files.shift();
}

export function listAttachedFiles(): string[] {
  return files.map((f) => f.name);
}

const fold = (s: string) =>
  String(s || "")
    .normalize("NFC")
    .trim()
    .toLowerCase();

/** Tìm tệp theo tên (không phân biệt hoa thường; cho phép thiếu đuôi .docx). Không nêu tên thì lấy tệp đính kèm gần nhất. */
export function findAttachedFile(name?: string): AttachedFile | null {
  if (!files.length) return null;
  const want = fold(name || "");
  if (!want) return files[files.length - 1];
  const exact = files.filter((f) => fold(f.name) === want || fold(f.name) === want + ".docx");
  if (exact.length) return exact[exact.length - 1];
  const partial = files.filter((f) => fold(f.name).includes(want));
  return partial.length === 1 ? partial[0] : null;
}

export function clearAttachedFiles(): void {
  files.length = 0;
}
