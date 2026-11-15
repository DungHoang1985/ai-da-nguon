import { checkConsistency } from "../src/core/consistency";
import { ParaRow } from "../src/core/outline";
const typed = (texts: string[]): ParaRow[] => texts.map((text, i) => ({ i, text }));
function run(name: string, texts: string[], checks?: any) {
  const r = checkConsistency(typed(texts), checks ? { checks } : {});
  console.log("== " + name);
  for (const x of r.issues) console.log("  ", x.kind, "¶" + x.p, x.message.slice(0, 220), "|", (x.excerpt||"").slice(0,100));
  if (!r.issues.length) console.log("   (none)");
}
const A = ["Điều 1. a"];
for (const s of [
 "Bên A góp 1/2 (một phần hai) vốn.",
 "Bên A góp 1/3 (một phần ba) vốn.",
 "Theo khoản 5 (một số trường hợp) của hợp đồng.",
 "Thanh toán 3 đợt (hai đợt đầu mỗi đợt 30%).",
 "Thời hạn hai mươi năm (20).",
 "Giá 100 triệu đồng (một trăm triệu đồng).",
 "Giá 1,5 tỷ đồng (một tỷ năm trăm triệu đồng).",
 "Giá 1.5 tỷ đồng (một tỷ rưỡi đồng).",
 "Giá 2,5 triệu (hai triệu rưỡi).",
 "Số tiền 450.000 đồng (bốn trăm năm mươi nghìn đồng).",
 "Diện tích 120,5 m2 (một trăm hai mươi phẩy năm mét vuông).",
 "Tổng 12 (mười hai) tháng; ngày 2 (hai) tháng 5 (năm) năm 2026.",
 "Nộp 5% (năm phần trăm) và 50% (năm mươi phần trăm).",
 "Phạt 0,05%/ngày (không phẩy không năm phần trăm một ngày).",
 "Đợt 1: 30% (ba mươi phần trăm); đợt 2: 70% (bảy mươi phần trăm).",
 "VND 1,000,000 (Bằng chữ: Một triệu đồng).",
 "1,000,000 VND (one million)",
 "Giá trị 1.000.000.000 VNĐ (Bằng chữ: Một tỷ đồng chẵn).",
 "Thời hạn 36 (ba mươi sáu) tháng, tức 3 (ba) năm.",
 "02 (hai) bản chính, mỗi bên giữ 01 (một) bản.",
 "Thực hiện trong 10 ngày (mười ngày) kể từ ngày nhận.",
 "Ký ngày 15/11/2026 (mười lăm tháng mười một năm hai nghìn không trăm hai mươi sáu).",
 "Số CMND 012345678 (không một hai ba bốn năm sáu bảy tám).",
 "5 (năm) lần; 6 (sáu) lần; 7 (bảy) lần; 4 (bốn) lần; 4 (tư) lần.",
 "Tầng 12A (mười hai A)",
 "Diện tích 1.200 m2 (một nghìn hai trăm mét vuông)",
 "Mã số 3 (ba) Ký hiệu A.",
]) run(s.slice(0,60), [...A, s], ["amount"]);
