/** Kho tạm trong bộ nhớ cho bản gốc của các tệp người dùng đính kèm (.docx để so sánh phiên bản, .pdf để đọc theo trang)
 *  — trước đây chỉ .docx: trong phiên làm việc này, để công cụ Word so sánh
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

export function unregisterAttachedFile(name: string): void {
  const at = files.findIndex((f) => f.name === name);
  if (at >= 0) files.splice(at, 1);
}

/** ext (ví dụ ".pdf") = chỉ lấy tệp có đuôi đó */
const ofExt = (ext?: string) => (ext ? files.filter((f) => f.name.toLowerCase().endsWith(ext)) : files);

export function listAttachedFiles(ext?: string): string[] {
  return ofExt(ext).map((f) => f.name);
}

const fold = (s: string) =>
  String(s || "")
    .normalize("NFC")
    .trim()
    .toLowerCase();

/** Tìm tệp theo tên (không phân biệt hoa thường; cho phép thiếu đuôi .docx). Không nêu tên thì lấy tệp đính kèm gần nhất. */
export function findAttachedFile(name?: string, ext?: string): AttachedFile | null {
  const files = ofExt(ext);
  if (!files.length) return null;
  const want = fold(name || "");
  if (!want) return files[files.length - 1];
  const exact = files.filter((f) => fold(f.name) === want || fold(f.name) === want + ".docx");
  if (exact.length) return exact[exact.length - 1];
  const partial = files.filter((f) => fold(f.name).includes(want));
  return partial.length === 1 ? partial[0] : null;
}

const IMAGE_EXT = /\.(png|jpe?g|gif|webp)$/i;

/** Ảnh người dùng đính kèm trong phiên này (để công cụ đặt ảnh vào slide) */
export function listAttachedImages(): string[] {
  return files.filter((f) => IMAGE_EXT.test(f.name)).map((f) => f.name);
}

/** Tìm ảnh đính kèm theo tên (không phân biệt hoa thường, cho phép thiếu đuôi hoặc gõ một phần tên nếu chỉ khớp một ảnh)
 *  hoặc theo số thứ tự 1, 2… trong danh sách. Không nêu thì lấy ảnh đính kèm gần nhất. */
export function findAttachedImage(ref?: string | number): AttachedFile | null {
  const imgs = files.filter((f) => IMAGE_EXT.test(f.name));
  if (!imgs.length) return null;
  const raw = ref === undefined || ref === null ? "" : String(ref).trim();
  if (!raw) return imgs[imgs.length - 1];
  if (/^\d+$/.test(raw)) return imgs[Number(raw) - 1] ?? null;
  const want = fold(raw);
  const exact = imgs.filter((f) => fold(f.name) === want || fold(f.name).replace(IMAGE_EXT, "") === want);
  if (exact.length) return exact[exact.length - 1];
  const partial = imgs.filter((f) => fold(f.name).includes(want));
  return partial.length === 1 ? partial[0] : null;
}

export function clearAttachedFiles(): void {
  files.length = 0;
}
