import { checkConsistency } from "../src/core/consistency";
import { ParaRow } from "../src/core/outline";
const typed = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
function run(name: string, texts: string[], checks?: any) {
  const r = checkConsistency(typed(texts), checks ? { checks } : {});
  console.log("== " + name);
  for (const x of r.issues) console.log("  ", x.kind, "¶" + x.p, x.message.slice(0, 200), "|", (x.excerpt||"").slice(0,100));
  if (!r.issues.length) console.log("   (none)");
}
const A = ["Điều 1. a","1.1. x","Điều 2. b","2.1. y"];
for (const s of [
 "Theo Điều 468 BLDS 2015 và Điều 12 LĐĐ 2024.",
 "Theo khoản 1 Điều 6 NĐ 99/2022/NĐ-CP, Điều 7 TT 01/2021/TT-BXD.",
 "Theo Điều 3 Luật số 27/2023/QH15.",
 "Bên B phải tuân thủ pháp luật và thực hiện đúng Điều 99 của Hợp đồng.",
 "Bên B phải tuân thủ quyết định của Bên A theo Điều 99 Hợp đồng.",
 "Điều 5 của Hợp đồng này, khoản 7 Điều 2, Điều 9",
 "Điều 1 và Điều 2 Điều lệ này; Section 5 of this Agreement; Article 12.",
 "the contract provides in Article 9 that",
 "Điều 1.1 và Điều 1.2; khoản 1.1 Điều 1; mục 2.1; mục 2.3",
 "Điều 02, Điều 001, Điều 1 .",
 "theo quy định tại Điều 1, Điều 2, khoản 5 Điều 3, điểm a khoản 1 Điều 2.",
 "Điều 2.1(a) và 2.1(b); Điều 2.1 (c)",
 "Tại Điều 7 Phụ lục 1, Điều 1 Phụ lục 2",
]) run(s.slice(0,70), [...A, s], ["reference"]);
