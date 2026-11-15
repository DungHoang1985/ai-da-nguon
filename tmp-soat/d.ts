import { checkConsistency } from "../src/core/consistency";
import { ParaRow } from "../src/core/outline";
const typed = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
function run(name: string, texts: string[], checks?: any) {
  const r = checkConsistency(typed(texts), checks ? { checks } : {});
  console.log("== " + name, "terms:", r.terms.join("|"), "parties:", r.parties.join("|"));
  for (const x of r.issues) console.log("  ", x.kind, "¶" + x.p, x.message.slice(0, 260), "|", (x.excerpt||"").slice(0,100));
  if (!r.issues.length) console.log("   (none)");
}
run("hd", [
 "Điều 1. Giải thích từ ngữ",
 "1.1. “Hợp đồng” là hợp đồng mua bán này.",
 "1.2. “Ngày làm việc” là ngày không phải thứ bảy, chủ nhật.",
 "Điều 2. Khác",
 "2.1. Bên A có quyền chấm dứt hợp đồng lao động với người lao động.",
 "2.2. Trong vòng 5 ngày làm việc, các hợp đồng thế chấp phải được đăng ký.",
 "2.3. Nếu hợp đồng bị vô hiệu thì Hợp đồng chấm dứt.",
], ["term","party"]);
// party variants
run("party", [
 "HỢP ĐỒNG",
 "BÊN BÁN (sau đây gọi là “Bên A”): Ông X",
 "BÊN MUA (sau đây gọi là “Bên B”): Bà Y",
 "Điều 1. Điều khoản",
 "1.1. Bên A bán cho Bên B. Bên Bán giao nhà. Bên thứ ba không được. Bên Mua trả tiền.",
 "1.2. Bên A hoặc Bên B có quyền. Bên Vay và Bên Cho Vay. Bên A Và B.",
 "1.3. Bên Nhận chuyển nhượng; Bên Có Quyền Lợi; Bên Còn Lại; Bên Được Bảo Lãnh.",
], ["term","party"]);
run("party2", [
 "CÔNG TY TNHH ABC",
 "BÊN A: Công ty ABC",
 "BÊN B: Ông Nguyễn",
 "Điều 1. X",
 "Bên A và Bên B thỏa thuận. Bên C không liên quan. Bên Thuê sử dụng. Bên Nào Vi Phạm sẽ bồi thường.",
 "Bên Thứ Ba. Ngân Hàng ABC; Hội Đồng Quản Trị; Giám Đốc Điều Hành; Ngày Làm Việc",
], ["term","party"]);
