/** PDF thử 2 trang dựng bằng mã cho bước Tự kiểm tra của công cụ đọc theo trang:
 *  trang 1 có lớp chữ (trên 300 ký tự), trang 2 chỉ có một hình chữ nhật xám (không có lớp chữ, như trang quét). */
export function testPdf(): string {
  const NL = "\n";
  const line = "MOC-TU-KIEM dong chu thu de kiem tra cong cu doc tep dinh kem theo trang cua add-in. ";
  const text = Array.from({ length: 6 }, (_, i) => `BT /F1 11 Tf 50 ${760 - i * 18} Td (${line}${i}) Tj ET`).join(NL);
  const scan = "0.8 g 80 300 400 300 re f";
  const stream = (body: string) => `<< /Length ${body.length} >>${NL}stream${NL}${body}${NL}endstream`;
  const objs = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 7 0 R >> >> >>",
    stream(text),
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 6 0 R /Resources << >> >>",
    stream(scan),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let out = "%PDF-1.4" + NL;
  const offs: number[] = [];
  objs.forEach((o, i) => {
    offs.push(out.length);
    out += `${i + 1} 0 obj${NL}${o}${NL}endobj${NL}`;
  });
  const xref = out.length;
  out += `xref${NL}0 ${objs.length + 1}${NL}0000000000 65535 f ${NL}` + offs.map((n) => `${String(n).padStart(10, "0")} 00000 n ${NL}`).join("");
  out += `trailer${NL}<< /Size ${objs.length + 1} /Root 1 0 R >>${NL}startxref${NL}${xref}${NL}%%EOF${NL}`;
  return out;
}
