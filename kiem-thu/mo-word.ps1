# Mở Word với một tệp thử kèm cổng gỡ lỗi WebView2 (để kiem-thu/cdp.mjs điều khiển khung add-in).
#   powershell -File kiem-thu/mo-word.ps1 -Tep <tệp .docx> [-Cong 9229]
#   powershell -File kiem-thu/mo-word.ps1 -Dong            chỉ đóng Word (không lưu) nếu mọi tài liệu đang mở đều là tệp thử
# An toàn: nếu Word, Excel hoặc PowerPoint đang mở tài liệu KHÔNG phải tệp thử thì dừng, không đóng gì (hỏi người dùng trước).
# Tệp thử = tệp nằm trong thư mục tạm của phiên hoặc có tên bắt đầu bằng "aidn-thu-" / "tham-do-".
param([string]$Tep = "", [int]$Cong = 9229, [switch]$Dong)
$ErrorActionPreference = 'Stop'

function La-Tep-Thu($name) { return ($name -match '^(aidn-thu-|tham-do-)') }

$khac = Get-Process excel, powerpnt -ErrorAction SilentlyContinue
if ($khac) { Write-Output "DỪNG: Excel hoặc PowerPoint đang chạy. Cần đóng trước (hỏi người dùng nếu đang có tài liệu)."; exit 2 }

if (Get-Process winword -ErrorAction SilentlyContinue) {
  $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
  $names = @($w.Documents | ForEach-Object { $_.Name })
  $la = @($names | Where-Object { -not (La-Tep-Thu $_) })
  if ($la.Count -gt 0) { Write-Output "DỪNG: Word đang mở tài liệu không phải tệp thử: $($la -join ', '). Không đóng."; exit 2 }
  foreach ($d in @($w.Documents)) { $d.Close(0) }   # 0 = không lưu
  if ($w.Documents.Count -ne 0) { Write-Output "DỪNG: còn tài liệu chưa đóng được."; exit 2 }
  $w.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
  for ($i = 0; $i -lt 20 -and (Get-Process winword -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
  if (Get-Process winword -ErrorAction SilentlyContinue) { Write-Output "DỪNG: Word chưa thoát hẳn."; exit 2 }
  Write-Output "Đã đóng Word (tệp thử: $($names -join ', '), không lưu)."
}
if ($Dong) { exit 0 }
if (-not $Tep) { Write-Output "Thiếu -Tep"; exit 1 }
$full = (Resolve-Path $Tep).Path
if (-not (La-Tep-Thu (Split-Path $full -Leaf))) { Write-Output "Tên tệp thử phải bắt đầu bằng aidn-thu- hoặc tham-do-"; exit 1 }

$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Cong --remote-allow-origins=*"
Start-Process winword.exe -ArgumentList "`"$full`""
for ($i = 0; $i -lt 60; $i++) {
  Start-Sleep -Seconds 1
  try {
    $list = (Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$Cong/json/list" -TimeoutSec 2).Content
    if ($list -match 'taskpane\.html') { Write-Output "Sẵn sàng: khung add-in đã mở sau $($i + 1) giây (cổng $Cong)."; exit 0 }
  } catch { }
}
Write-Output "Chưa thấy khung add-in sau 60 giây."
exit 1
