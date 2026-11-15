import { checkConsistency } from "../src/core/consistency";
import { buildOutline, ParaRow } from "../src/core/outline";
function run(name: string, rows: ParaRow[], checks?: any) {
  const r = checkConsistency(rows, checks ? { checks } : {});
  console.log("== " + name);
  for (const x of r.issues) console.log("  ", x.kind, "¶" + x.p, x.message.slice(0, 230), "|", (x.excerpt||"").slice(0,100));
  if (!r.issues.length) console.log("   (none)");
}
const T = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
run("typed1", T([
 "Điều 1. A", "1. x", "2. y", "3. z", "Điều 2. B", "1. x", "a) p", "b) q", "2. y", "a) p", "b) q", "c) r", "đ) s", "Điều 3. C", "1. x", "2. y",
 "Điều 4. D", "1. x", "a) p", "b) q", "d) r", "Điều 5", "1. x", "3. y", "Điều 6. E", "a) k",
]), ["numbering"]);
// Word list: multiple lists; recitals list "1." "2." typed before Điều
const L = (i: number, text: string, listString: string, listLevel: number, listId: number, extra: Partial<ParaRow> = {}): ParaRow => ({ i, text, listString, listLevel, listId, ...extra });
run("lists", [
 L(0, "Căn cứ Bộ luật Dân sự;", "1.", 0, 5),
 L(1, "Căn cứ Luật Nhà ở;", "2.", 0, 5),
 L(2, "Điều 1. A", "Điều 1.", 0, 7),
 L(3, "x", "1.1.", 1, 7),
 L(4, "x", "1.2.", 1, 7),
 L(5, "x", "a)", 2, 7),
 L(6, "x", "b)", 2, 7),
 L(7, "Điều 2. B", "Điều 2.", 0, 7),
 L(8, "x", "2.1.", 1, 7),
 L(9, "x", "(i)", 2, 7),
 L(10, "x", "(ii)", 2, 7),
 L(11, "Điều 3. C", "Điều 3.", 0, 7),
 L(12, "Bên A nhận", "•", 1, 9),
 L(13, "Bên B nhận", "•", 1, 9),
 L(14, "x", "3.1.", 1, 7),
 L(15, "Phụ lục 1", "", 0, undefined as any, {style:"Heading1"}),
 L(16, "Điều 1. X", "Điều 1.", 0, 11),
 L(17, "Điều 2. Y", "Điều 2.", 0, 11),
], ["numbering","reference","term"]);
// deleted
run("deleted", [
 L(0, "Điều 1. A", "Điều 1.", 0, 7),
 L(1, "x", "1.1.", 1, 7),
 L(2, "", "1.2.", 1, 7, {deleted: true}),
 L(3, "x", "1.3.", 1, 7),
 L(4, "Điều 2. B", "Điều 2.", 0, 7),
 L(5, "", "Điều 3.", 0, 7, {deleted: true}),
 L(6, "Điều 4. D", "Điều 4.", 0, 7),
 L(7, "Điều 5. D", "Điều 5.", 0, 7),
 L(8, "x", "1.", 1, 7),
 L(9, "x", "2.", 1, 7),
], ["numbering","reference"]);
console.log(buildOutline([
 L(0, "Điều 1. A", "Điều 1.", 0, 7),
 L(1, "x", "1.", 1, 7),
 L(2, "x", "2.", 1, 7)]).map(o=>[o.i,o.num,o.key,o.kind].join(",")));
