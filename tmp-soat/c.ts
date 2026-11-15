import { checkConsistency } from "../src/core/consistency";
import { ParaRow } from "../src/core/outline";
const typed = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
function run(name: string, texts: string[], checks?: any) {
  const r = checkConsistency(typed(texts), checks ? { checks } : {});
  console.log("== " + name);
  for (const x of r.issues) console.log("  ", x.kind, "¶" + x.p, x.message.slice(0, 230), "|", (x.excerpt||"").slice(0,100));
  if (!r.issues.length) console.log("   (none)");
}
const base = [
 "Điều 1. Giải thích từ ngữ",
 "1. “Hợp Đồng” là hợp đồng này và các phụ lục.",
 "2. “Ngày Bàn Giao” là ngày Bên A bàn giao nhà.",
 "a) điểm a",
 "b) điểm b",
 "c) điểm c",
 "d) điểm d",
 "đ) điểm đ",
 "e) điểm e",
 "Điều 2. Thanh toán",
 "1. Bên B thanh toán.",
 "2. Thanh toán theo khoản 1 Điều này và điểm đ khoản 2 Điều 1.",
 "3. Theo điểm a, b, c khoản 2 Điều 1; Điều 1 và Điều 2.",
 "Điều 3. Khác",
 "1. Căn cứ Điều 3 Luật Nhà ở 2023 và các Điều 5, 6 và 7 Luật Kinh doanh bất động sản.",
 "2. Theo quy định tại Điều 100 của Bộ luật Dân sự và khoản 3 Điều 200 Nghị định này.",
 "3. Xem Điều 15 Hợp đồng chuyển nhượng số 12/2020; Điều 20 Phụ lục 01.",
 "4. Điều 5 của Thông tư 12.",
 "5. Điều kiện thanh toán xem Điều 2.",
 "6. Điểm 1 Điều 2 và mục 3.2 và Phụ lục I, Chương V.",
]
run("refs", base, ["reference"]);
run("terms", base, ["term","party"]);
run("num", base, ["numbering"]);
