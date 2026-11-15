# Mở Word, Excel hoặc PowerPoint với một tệp thử kèm cổng gỡ lỗi WebView2 (để kiem-thu/cdp.mjs điều khiển khung add-in).
#   powershell -File kiem-thu/mo-office.ps1 -Tep <tệp thử> [-Cong 9229]      mở tệp (ứng dụng theo đuôi tệp: .docx, .xlsx, .pptx)
#   powershell -File kiem-thu/mo-office.ps1 -Dong                            chỉ đóng các ứng dụng Office đang mở tệp thử (không lưu)
# An toàn: nếu Word, Excel hoặc PowerPoint đang mở tài liệu KHÔNG phải tệp thử thì dừng, không đóng gì (hỏi người dùng trước).
# Tệp thử = tệp có tên bắt đầu bằng "aidn-thu-" hoặc "tham-do-". Mọi ứng dụng Office phải đóng hết trước khi mở
# (WebView2 chỉ nhận cổng gỡ lỗi lúc khởi động; ba ứng dụng dùng chung một cổng).
param([string]$Tep = "", [int]$Cong = 9229, [switch]$Dong)
$ErrorActionPreference = 'Stop'

function La-Tep-Thu($name) { return ($name -match '^(aidn-thu-|tham-do-)') }
function Cho-Thoat($proc) {
  for ($i = 0; $i -lt 40 -and (Get-Process $proc -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
  return -not (Get-Process $proc -ErrorAction SilentlyContinue)
}

# Word
if (Get-Process winword -ErrorAction SilentlyContinue) {
  $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
  $names = @($w.Documents | ForEach-Object { $_.Name })
  # Tài liệu ẩn (không có cửa sổ, không tên) do công cụ so sánh phiên bản dựng từ tệp đính kèm: không phải tài liệu của người dùng
  $la = @($w.Documents | Where-Object { $_.Windows.Count -gt 0 -and -not (La-Tep-Thu $_.Name) } | ForEach-Object { $_.Name })
  if ($la.Count -gt 0) { Write-Output "DUNG: Word dang mo tai lieu khong phai tep thu: $($la -join ', '). Khong dong."; exit 2 }
  foreach ($d in @($w.Documents)) { $d.Close(0) }   # 0 = không lưu
  if ($w.Documents.Count -ne 0) { Write-Output "DUNG: Word con tai lieu chua dong duoc."; exit 2 }
  $w.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
  if (-not (Cho-Thoat winword)) { Write-Output "DUNG: Word chua thoat han."; exit 2 }
  Write-Output "Da dong Word (tep thu: $($names -join ', '), khong luu)."
}
# Excel
if (Get-Process excel -ErrorAction SilentlyContinue) {
  $x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  $names = @($x.Workbooks | ForEach-Object { $_.Name })
  $la = @($names | Where-Object { -not (La-Tep-Thu $_) })
  if ($la.Count -gt 0) { Write-Output "DUNG: Excel dang mo so tinh khong phai tep thu: $($la -join ', '). Khong dong."; exit 2 }
  $x.DisplayAlerts = $false
  foreach ($b in @($x.Workbooks)) { $b.Close($false) }
  if ($x.Workbooks.Count -ne 0) { Write-Output "DUNG: Excel con so tinh chua dong duoc."; exit 2 }
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  # Excel sau Quit hay còn một tiến trình ẩn: đã chắc không còn sổ tính nào thì mới kết thúc tiến trình
  if (-not (Cho-Thoat excel)) { Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force; [void](Cho-Thoat excel) }
  Write-Output "Da dong Excel (tep thu: $($names -join ', '), khong luu)."
}
# PowerPoint
if (Get-Process powerpnt -ErrorAction SilentlyContinue) {
  $p = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
  $names = @($p.Presentations | ForEach-Object { $_.Name })
  $la = @($names | Where-Object { -not (La-Tep-Thu $_) })
  if ($la.Count -gt 0) { Write-Output "DUNG: PowerPoint dang mo bai khong phai tep thu: $($la -join ', '). Khong dong."; exit 2 }
  foreach ($d in @($p.Presentations)) { $d.Saved = -1; $d.Close() }
  if ($p.Presentations.Count -ne 0) { Write-Output "DUNG: PowerPoint con bai chua dong duoc."; exit 2 }
  $p.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)
  if (-not (Cho-Thoat powerpnt)) { Write-Output "DUNG: PowerPoint chua thoat han."; exit 2 }
  Write-Output "Da dong PowerPoint (tep thu: $($names -join ', '), khong luu)."
}
if ($Dong) { exit 0 }
if (-not $Tep) { Write-Output "Thieu -Tep"; exit 1 }
$full = (Resolve-Path $Tep).Path
if (-not (La-Tep-Thu (Split-Path $full -Leaf))) { Write-Output "Ten tep thu phai bat dau bang aidn-thu- hoac tham-do-"; exit 1 }
$exe = switch -Regex ($full) { '\.docx$' { 'winword.exe' } '\.xlsx$' { 'excel.exe' } '\.pptx$' { 'powerpnt.exe' } default { '' } }
if (-not $exe) { Write-Output "Tep phai la .docx, .xlsx hoac .pptx"; exit 1 }

$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Cong --remote-allow-origins=*"
Start-Process $exe -ArgumentList "`"$full`""
for ($i = 0; $i -lt 90; $i++) {
  Start-Sleep -Seconds 1
  try {
    $list = (Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$Cong/json/list" -TimeoutSec 2).Content
    if ($list -match 'taskpane\.html') { Write-Output "San sang: khung add-in da mo sau $($i + 1) giay (cong $Cong)."; exit 0 }
  } catch { }
}
Write-Output "Chua thay khung add-in sau 90 giay."
exit 1
