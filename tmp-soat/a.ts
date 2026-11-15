import { checkConsistency, parseVnNumber, parseFigure } from "../src/core/consistency";
import { ParaRow } from "../src/core/outline";
const typed = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
function run(name: string, texts: string[], checks?: any) {
  const r = checkConsistency(typed(texts), checks ? { checks } : {});
  console.log("== " + name);
  for (const x of r.issues) console.log("  ", x.kind, "¶" + x.p, x.message.slice(0, 200), "|", x.excerpt?.slice(0,100));
  if (!r.issues.length) console.log("   (none)");
}
// clean contract
run("clean1", [
 "HỢP ĐỒNG THUÊ NHÀ",
 "Hôm nay, ngày 15 tháng 11 năm 2026, tại Hà Nội, chúng tôi gồm:",
 "BÊN CHO THUÊ (sau đây gọi là “Bên A”): Ông Nguyễn Văn A",
 "BÊN THUÊ (sau đây gọi là “Bên B”): Bà Trần Thị B",
 "Điều 1. Đối tượng hợp đồng",
 "1.1. Bên A cho Bên B thuê căn nhà theo quy định tại Điều 472 Bộ luật Dân sự 2015.",
 "1.2. Thời hạn thuê là 05 (năm) năm kể từ ngày ký.",
 "1.3. Giá thuê là 15.000.000 đồng (mười lăm triệu đồng) mỗi tháng, theo khoản 1 Điều này.",
 "Điều 2. Quyền và nghĩa vụ",
 "2.1. Bên A có quyền: a) nhận tiền; b) kiểm tra; c) đơn phương chấm dứt theo Điều 3 của Hợp đồng này.",
 "2.2. Bên B thanh toán theo điểm a khoản 1 Điều 1.",
 "Điều 3. Chấm dứt",
 "3.1. Theo khoản 2 Điều 2 của hợp đồng, v.v...",
 "3.2. Phạt 8,5% giá trị (tám phẩy năm phần trăm).",
 "3.3. Giá trị 1.500.000.000 đồng (một tỷ năm trăm triệu đồng).",
 "3.4. Phạt vi phạm 10% (mười phần trăm) giá trị.",
 "3.5. Thời hạn 30 (ba mươi) ngày; trả trước 3 (ba) tháng.",
 "3.6. Theo Điều 10 Nghị định 99/2022/NĐ-CP và khoản 3 Điều 5 Luật Đất đai.",
]);
run("amt", [
 "Điều 1. a", "1.1. Giá 2.500.000 đồng (hai triệu rưỡi đồng).",
 "1.2. Giá 1.250.000 đồng (một triệu hai trăm năm mươi nghìn đồng).",
 "1.3. 105 (một trăm linh năm) ngày.",
 "1.4. 1.000.000.000 đồng (một tỷ đồng chẵn).",
 "1.5. 15 (mười lăm) ngày; 25 (hai mươi lăm); 21 (hai mươi mốt); 24 (hai mươi tư); 35 (ba mươi lăm); 101 (một trăm lẻ một)",
 "1.6. 2.000 đồng (hai nghìn đồng); 12.345 (mười hai nghìn ba trăm bốn mươi lăm)",
 "1.7. 1,5 (một phẩy năm); 2 (hai) bên; 3 (ba) năm; 20 (hai mươi) năm; 05 (năm) năm; 1.200 (một nghìn hai trăm)",
 "1.8. 7 (bảy) ngày làm việc; 15 (mười lăm) ngày làm việc kể từ ngày 15 tháng 11 năm 2026",
 "1.9. 5 (năm) năm; 5 năm (năm năm); 50 (năm mươi); 5.000 (năm nghìn); 2.000.000 (hai triệu)",
 "1.10. 1.100.000 đồng (một triệu một trăm nghìn đồng), 1.010.000 (một triệu không trăm mười nghìn)",
 "1.11. 1.001.000 đồng (một triệu không trăm linh một nghìn); 2.050.000 (hai triệu không trăm năm mươi nghìn)",
]);
console.log(parseVnNumber("một triệu không trăm mười nghìn".split(" ")), parseVnNumber("hai triệu rưỡi".split(" ")), parseVnNumber("một nghìn hai trăm".split(" ")), parseVnNumber("một triệu hai trăm năm mươi nghìn".split(" ")));
console.log(parseVnNumber("một tỷ hai trăm triệu".split(" ")), parseVnNumber("hai tỷ".split(" ")),parseVnNumber("một nghìn tỷ".split(" ")),parseVnNumber("năm trăm nghìn".split(" ")), parseVnNumber("ba trăm hai mươi".split(" ")));
console.log(parseVnNumber("hai mươi mốt".split(" ")), parseVnNumber("một trăm rưỡi".split(" ")),parseVnNumber("mười hai".split(" ")),parseVnNumber("một trăm lẻ năm nghìn".split(" ")), parseVnNumber("tám mươi tư triệu".split(" ")));
console.log(parseFigure("1.500"), parseFigure("1,5"), parseFigure("15,000"), parseFigure("1.5"),parseFigure("2.50"), parseFigure("0,5"));
