/** Ngưỡng luật hay gặp mà model miễn phí chép nguyên số người dùng đưa dù trái luật (vòng thử 3, 24/11/2026:
 *  cọc 10% của chủ đầu tư nhà hình thành trong tương lai; thử việc 90 ngày, lương thử việc 70% với kỹ sư đại học).
 *  Chỉ báo (`check_now`), không sửa; ưu tiên không báo nhầm: mỗi luật chỉ xét khi văn bản có đúng ngữ cảnh.
 *  Theo quy tắc của kho: không ghi số hiệu văn bản, chỉ tên và năm. Soát lại khi luật đổi. */

const pct = (s: string) => Number(s.replace(",", "."));
/** Số (1–3 chữ số, có thể có phần thập phân) không bị cắt giữa chừng: "100%" không thành "00%" (soát mã 4.11.11, #1) */
const NUM = "(?<![\\d.,])(\\d{1,3}(?:[,.]\\d+)?)";

/** Bên bán / bên nhận cọc là doanh nghiệp (dòng giới thiệu bên có "công ty", "tập đoàn"…) */
const ORG_PARTY = /^[#*\s]*BÊN[^:\n]{0,60}:\s*[*_]*\s*(?:CÔNG TY|TỔNG CÔNG TY|TẬP ĐOÀN|DOANH NGHIỆP)/imu;

/** Tiền đặt cọc chủ đầu tư thu khi bán nhà ở, công trình hình thành trong tương lai: không quá 5% giá bán (Luật Kinh doanh bất động sản năm 2023).
 *  Chỉ xét khi: có đặt cọc; có chủ đầu tư hoặc nhà hình thành trong tương lai; một bên là doanh nghiệp; con số là tỷ lệ của GIÁ (bán, mua, thuê mua)
 *  và không phải mức phạt cọc (soát mã 4.11.11, #4). */
function depositCap(text: string): string | null {
  if (!/đặt cọc/iu.test(text) || !/chủ đầu tư|hình thành trong tương lai/iu.test(text) || !ORG_PARTY.test(text)) return null;
  const re = new RegExp(`(phạt\\s+)?(?:tiền\\s+)?(?:đặt\\s+)?cọc[^\\n.;]{0,80}?${NUM}\\s*%\\s*(?:của\\s+|trên\\s+|tổng\\s+)?(?:giá|tổng giá)\\s*(?:bán|mua|thuê mua|trị)`, "giu");
  for (const m of text.matchAll(re)) {
    if (m[1]) continue;
    const v = pct(m[2]);
    if (v > 5 && v <= 100)
      return `Tiền đặt cọc ${m[2]}% giá bán: theo Luật Kinh doanh bất động sản năm 2023, chủ đầu tư chỉ được thu tiền đặt cọc không quá 5% giá bán, giá thuê mua nhà ở, công trình (khi đã đủ điều kiện đưa vào kinh doanh) — hạ mức cọc về tối đa 5% hoặc nêu rõ rủi ro với người dùng, không chép nguyên`;
  }
  return null;
}

/** Thử việc (Bộ luật Lao động năm 2019): tối đa 180 ngày với người quản lý doanh nghiệp, 60 ngày với công việc cần trình độ
 *  cao đẳng trở lên, 30 ngày trung cấp, 6 ngày làm việc với việc khác; lương thử việc ít nhất 85% lương của công việc đó. */
function probation(text: string): string[] {
  if (!/thử việc/iu.test(text)) return [];
  const out: string[] = [];
  // Người quản lý doanh nghiệp: xét dòng nêu chức danh, vị trí, công việc (không xét dòng người ký phía công ty)
  const roleLines = text.split("\n").filter((l) => /chức danh|chức vụ|vị trí|công việc|đảm nhiệm|giữ chức/iu.test(l) && !/^[#*\s]*đại diện/iu.test(l));
  const manager = /quản lý doanh nghiệp|giám đốc|kế toán trưởng/iu.test(roleLines.join("\n"));
  const high = /đại học|cao đẳng|kỹ sư|cử nhân|thạc sĩ|tiến sĩ/iu.test(text);
  for (const m of text.matchAll(new RegExp(`thử việc[^\\n.;]{0,60}?${NUM}\\s*ngày`, "giu"))) {
    const d = Number(m[1]);
    const cap = manager ? 180 : high ? 60 : 0;
    if (d > 180 || (cap && d > cap)) {
      out.push(`Thử việc ${d} ngày: Bộ luật Lao động năm 2019 cho tối đa ${d > 180 ? "180 ngày (chỉ với người quản lý doanh nghiệp)" : `${cap} ngày với công việc cần trình độ ${manager ? "người quản lý doanh nghiệp" : "cao đẳng trở lên"}`} — sửa thời gian thử việc hoặc nêu rõ với người dùng`);
      break;
    }
  }
  // Tỷ lệ phải là tỷ lệ của LƯƠNG ("70% mức lương", "lương thử việc: 70%"), không phải phụ cấp, thưởng đứng sau (soát mã 4.11.11, #5)
  const re = new RegExp(`(?:lương|tiền lương)\\s+(?:trong\\s+(?:thời gian\\s+)?)?thử việc(?:[^\\n.;%]{0,40}?)${NUM}\\s*%(\\s*(?:của\\s+|so với\\s+)?(?:mức\\s+)?(?:lương|tiền lương))?|thử việc[^\\n.;%]{0,40}?${NUM}\\s*%\\s*(?:của\\s+|so với\\s+)?(?:mức\\s+)?(?:lương|tiền lương)`, "giu");
  for (const m of text.matchAll(re)) {
    const raw = m[1] || m[3];
    const between = m[0].slice(0, m[0].indexOf(raw));
    if (/phụ cấp|thưởng|ăn trưa|xăng xe|điện thoại/iu.test(between)) continue;
    const v = pct(raw);
    if (v < 85) {
      out.push(`Lương thử việc ${raw}%: Bộ luật Lao động năm 2019 buộc ít nhất 85% mức lương của công việc đó — sửa mức lương thử việc hoặc nêu rõ với người dùng`);
      break;
    }
  }
  return out;
}

/** Hợp đồng lao động xác định thời hạn: không quá 36 tháng (Bộ luật Lao động năm 2019). Chỉ xét câu nêu rõ "xác định thời hạn" (soát mã 4.11.11, #12) */
function fixedTerm(text: string): string | null {
  if (!/hợp đồng lao động/iu.test(text)) return null;
  const m = new RegExp(`xác định thời hạn[^\\n.;]{0,40}?${NUM}\\s*tháng`, "iu").exec(text);
  const n = m ? Number(m[1]) : 0;
  return n > 36 ? `Hợp đồng lao động xác định thời hạn ${n} tháng: Bộ luật Lao động năm 2019 cho tối đa 36 tháng — sửa hoặc nêu rõ với người dùng` : null;
}

export function legalThresholds(lines: string[]): string[] {
  const text = lines.join("\n").normalize("NFC");
  return [depositCap(text), ...probation(text), fixedTerm(text)].filter((x): x is string => !!x);
}
