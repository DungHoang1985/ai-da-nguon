/** Đọc mục lục điều khoản của một tệp .docx bằng đúng mã của add-in (Đợt 12, phần D): in số đoạn, số điều và vài mục đầu.
 *  Chạy: TS_NODE_COMPILER_OPTIONS='{"module":"commonjs","moduleResolution":"node"}' npx ts-node --transpile-only kiem-thu/doc-dieu-docx.ts <tệp.docx> [điều cần tra] */
import * as fs from "fs";
import JSZip from "jszip";
import { docxParagraphs, paraDisplay } from "../src/core/docxparas";
import { buildOutline, findClause } from "../src/core/outline";

(async () => {
  const zip = await JSZip.loadAsync(fs.readFileSync(process.argv[2]));
  const get = async (n: string) => (await zip.file(n)?.async("string")) || undefined;
  const rows = docxParagraphs({ document: (await get("word/document.xml"))!, styles: await get("word/styles.xml"), numbering: await get("word/numbering.xml") });
  const items = buildOutline(rows);
  console.log(`${rows.length} đoạn, ${rows.filter((r) => r.listString).length} đoạn có số, ${items.length} mục`);
  for (const it of items.slice(0, 12)) console.log(`  ${it.num} [¶${it.i}–${it.end}] ${it.title.slice(0, 50)}`);
  const q = process.argv[3];
  if (q) for (const it of findClause(items, q)) console.log(`>> ${q}: ¶${it.i}–${it.end}\n` + rows.slice(it.i, it.end + 1).map(paraDisplay).join("\n").slice(0, 800));
})();
