// Dựng một tệp PDF "quét" 3 trang cho Đợt 1B: mỗi trang là MỘT ẢNH (không có lớp chữ), chữ hơi nghiêng và có nhiễu như bản quét.
// Nội dung bịa để thử (biên bản nghiệm thu, số liệu không có thật). Ba số liệu để hỏi lại nằm ở SO_LIEU.
//   node kiem-thu/pdf-quet-thu.mjs <tệp ra.pdf>
// Cần Windows (dùng System.Drawing của PowerShell để vẽ chữ thành ảnh JPEG).
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";

export const SO_LIEU = [
  { hoi: "Giá trị hợp đồng", dung: "48.750.000.000" },
  { hoi: "Diện tích sàn bàn giao", dung: "12.860" },
  { hoi: "Số tiền giữ lại bảo hành", dung: "2.437.500.000" },
];

const TRANG = [
  [
    "CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN HỒNG LĨNH",
    "BIÊN BẢN NGHIỆM THU HOÀN THÀNH VÀ BÀN GIAO CÔNG TRÌNH",
    "(bản thử của AI Đa Nguồn, nội dung không có thật)",
    "",
    "Số: 37/2026/BB-NT                    Ngày 28 tháng 10 năm 2026",
    "",
    "1. Công trình: Nhà xưởng số 2 thuộc Khu công nghiệp Tân Phú.",
    "2. Chủ đầu tư: Công ty Cổ phần Bất động sản Hồng Lĩnh.",
    "3. Nhà thầu thi công: Công ty TNHH Xây dựng Lam Giang.",
    "4. Hợp đồng thi công số 12/2025/HĐXD ký ngày 15 tháng 3 năm 2025.",
    "5. Giá trị hợp đồng: 48.750.000.000 đồng (đã gồm thuế giá trị gia tăng).",
    "6. Thời gian thi công thực tế: từ ngày 01/4/2025 đến ngày 20/10/2026.",
    "",
    "Hai bên cùng tiến hành nghiệm thu với các nội dung ghi ở trang sau.",
  ],
  [
    "NỘI DUNG NGHIỆM THU",
    "",
    "1. Khối lượng: nhà thầu đã thi công xong toàn bộ hạng mục theo thiết kế.",
    "2. Diện tích sàn bàn giao: 12.860 m2, gồm xưởng sản xuất và kho.",
    "3. Chiều cao thông thủy: 9,5 mét; tải trọng nền: 3 tấn trên mỗi mét vuông.",
    "4. Hệ thống phòng cháy chữa cháy đã được nghiệm thu ngày 12/10/2026.",
    "5. Chất lượng: đạt yêu cầu của hồ sơ thiết kế và tiêu chuẩn áp dụng.",
    "6. Tồn tại: sơn lại 02 vị trí tường trục C; thay 03 bộ đèn chiếu sáng.",
    "   Thời hạn khắc phục: trước ngày 15 tháng 11 năm 2026.",
    "",
    "Kết luận: chấp nhận nghiệm thu hoàn thành công trình để đưa vào sử dụng.",
  ],
  [
    "THANH TOÁN VÀ BẢO HÀNH",
    "",
    "1. Giá trị đã thanh toán đến ngày nghiệm thu: 41.437.500.000 đồng.",
    "2. Số tiền giữ lại bảo hành: 2.437.500.000 đồng (bằng 5% giá trị hợp đồng).",
    "3. Giá trị còn phải thanh toán đợt cuối: 4.875.000.000 đồng.",
    "4. Thời hạn bảo hành công trình: 24 tháng kể từ ngày ký biên bản này.",
    "5. Tiền giữ lại được hoàn trả trong 15 ngày sau khi hết thời hạn bảo hành.",
    "",
    "Biên bản lập thành 04 bản có giá trị như nhau, mỗi bên giữ 02 bản.",
    "",
    "ĐẠI DIỆN CHỦ ĐẦU TƯ                    ĐẠI DIỆN NHÀ THẦU",
    "(đã ký và đóng dấu)                        (đã ký và đóng dấu)",
  ],
];

const PS = String.raw`
param([string]$Dir, [string]$Json)
Add-Type -AssemblyName System.Drawing
$pages = Get-Content -Raw -Encoding UTF8 $Json | ConvertFrom-Json
$rnd = New-Object System.Random 20261115
$n = 0
foreach ($lines in $pages) {
  $n++
  $bmp = New-Object System.Drawing.Bitmap 1240, 1754
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::FromArgb(246, 244, 238))
  $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
  # Bản quét thường hơi nghiêng
  $g.TranslateTransform(620, 877)
  $g.RotateTransform((($n % 2) * 1.4) - 0.7)
  $g.TranslateTransform(-620, -877)
  $font = New-Object System.Drawing.Font 'Times New Roman', 21
  $bold = New-Object System.Drawing.Font 'Times New Roman', 23, ([System.Drawing.FontStyle]::Bold)
  $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(38, 38, 44))
  $y = 150
  $k = 0
  foreach ($line in $lines) {
    $f = $font
    if ($k -lt 2 -and $line -cmatch '^[^a-z]+$') { $f = $bold }
    $g.DrawString([string]$line, $f, $brush, 110, $y)
    $y += 62
    $k++
  }
  $g.ResetTransform()
  # Nhiễu hạt và vệt mờ của máy quét
  for ($i = 0; $i -lt 2600; $i++) {
    $c = 150 + $rnd.Next(90)
    $bmp.SetPixel($rnd.Next(1240), $rnd.Next(1754), [System.Drawing.Color]::FromArgb($c, $c, $c))
  }
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(40, 120, 120, 120)), 3
  $g.DrawLine($pen, 0, 40 + $rnd.Next(60), 1240, 30 + $rnd.Next(60))
  $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
  $ep = New-Object System.Drawing.Imaging.EncoderParameters 1
  $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), 62L
  $bmp.Save((Join-Path $Dir ("trang-$n.jpg")), $codec, $ep)
  $g.Dispose(); $bmp.Dispose()
}
Write-Output "da ve $n trang"
`;

/** PDF tối giản: mỗi trang A4 chứa đúng một ảnh JPEG phủ kín trang (bộ lọc DCTDecode), không có đối tượng chữ nào */
export function pdfTuAnh(jpegs) {
  const objs = [];
  const add = (body) => {
    objs.push(body);
    return objs.length;
  };
  const catalog = add(null);
  const pagesId = add(null);
  const kids = [];
  for (const jpg of jpegs) {
    // Kích thước ảnh nằm trong đoạn SOF của JPEG
    let w = 1240;
    let h = 1754;
    for (let i = 2; i < jpg.length - 9; ) {
      if (jpg[i] !== 0xff) break;
      const marker = jpg[i + 1];
      const len = jpg.readUInt16BE(i + 2);
      if (marker >= 0xc0 && marker <= 0xc3) {
        h = jpg.readUInt16BE(i + 5);
        w = jpg.readUInt16BE(i + 7);
        break;
      }
      i += 2 + len;
    }
    const img = add({ dict: `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpg.length} >>`, stream: jpg });
    const content = Buffer.from(`q 595.28 0 0 841.89 0 0 cm /Im0 Do Q`);
    const cont = add({ dict: `<< /Length ${content.length} >>`, stream: content });
    kids.push(add(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595.28 841.89] /Resources << /XObject << /Im0 ${img} 0 R >> >> /Contents ${cont} 0 R >>`));
  }
  objs[catalog - 1] = `<< /Type /Catalog /Pages ${pagesId} 0 R >>`;
  objs[pagesId - 1] = `<< /Type /Pages /Kids [${kids.map((k) => `${k} 0 R`).join(" ")}] /Count ${kids.length} >>`;
  const chunks = [Buffer.from("%PDF-1.4\n%\xe2\xe3\xcf\xd3\n", "latin1")];
  const offsets = [];
  let pos = chunks[0].length;
  objs.forEach((o, k) => {
    offsets.push(pos);
    const parts = typeof o === "string" ? [Buffer.from(`${k + 1} 0 obj\n${o}\nendobj\n`, "latin1")] : [Buffer.from(`${k + 1} 0 obj\n${o.dict}\nstream\n`, "latin1"), o.stream, Buffer.from("\nendstream\nendobj\n", "latin1")];
    for (const p of parts) {
      chunks.push(p);
      pos += p.length;
    }
  });
  const xref = `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map((o) => String(o).padStart(10, "0") + " 00000 n \n").join("")}trailer\n<< /Size ${objs.length + 1} /Root ${catalog} 0 R >>\nstartxref\n${pos}\n%%EOF\n`;
  chunks.push(Buffer.from(xref, "latin1"));
  return Buffer.concat(chunks);
}

const out = process.argv[2];
if (!out) {
  console.log("Cách dùng: node kiem-thu/pdf-quet-thu.mjs <tệp ra.pdf>");
  process.exit(1);
}
const dir = fs.mkdtempSync(path.join(os.tmpdir(), "aidn-pdf-"));
const ps1 = path.join(dir, "ve.ps1");
// Tệp .ps1 phải có BOM để PowerShell 5.1 đọc đúng UTF-8
fs.writeFileSync(ps1, "﻿" + PS, "utf8");
fs.writeFileSync(path.join(dir, "trang.json"), JSON.stringify(TRANG), "utf8");
console.log(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, "-Dir", dir, "-Json", path.join(dir, "trang.json")], { encoding: "utf8" }).trim());
const jpegs = TRANG.map((_, k) => fs.readFileSync(path.join(dir, `trang-${k + 1}.jpg`)));
fs.writeFileSync(out, pdfTuAnh(jpegs));
fs.rmSync(dir, { recursive: true, force: true });
console.log(`Đã tạo ${out} (${TRANG.length} trang, ${(fs.statSync(out).size / 1024).toFixed(0)} KB, không có lớp chữ)`);
