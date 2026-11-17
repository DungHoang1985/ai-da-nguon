# Chạy một chương trình trên DESKTOP RIÊNG (desktop Win32 ẩn, cùng phiên Windows) để cửa sổ của nó không hiện lên
# màn hình anh Dũng đang làm việc, không giành bàn phím, không đụng chuột.
#   powershell -File kiem-thu/desktop-rieng.ps1 -Lenh <exe> [-ThamSo "<tham số>"] [-Ten aidn-thu]   chạy, in PID
#   powershell -File kiem-thu/desktop-rieng.ps1 -LietKe [-Ten aidn-thu]                              liệt kê cửa sổ trên desktop đó
# Biến môi trường của tiến trình gọi (ví dụ WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS) được truyền sang chương trình.
# Desktop tự mất khi không còn tiến trình nào trên đó.
param([string]$Lenh = "", [string]$ThamSo = "", [string]$Ten = "aidn-thu", [switch]$LietKe)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class DesktopRieng {
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] struct STARTUPINFO { public int cb; public string lpReserved; public string lpDesktop; public string lpTitle;
    public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags; public short wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError; }
  [StructLayout(LayoutKind.Sequential)] struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int dwProcessId, dwThreadId; }
  [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr CreateDesktop(string name, IntPtr dev, IntPtr devmode, int flags, uint access, IntPtr sa);
  [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr OpenDesktop(string name, int flags, bool inherit, uint access);
  [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr h);
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern bool CreateProcess(string app, StringBuilder cmd, IntPtr pa, IntPtr ta, bool inherit, int flags, IntPtr env, string cwd, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
  [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
  delegate bool EnumProc(IntPtr hwnd, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumDesktopWindows(IntPtr desk, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  const uint ALL = 0x10000000; // GENERIC_ALL
  public static int Run(string desk, string cmdline) {
    IntPtr d = CreateDesktop(desk, IntPtr.Zero, IntPtr.Zero, 0, ALL, IntPtr.Zero);
    if (d == IntPtr.Zero) throw new Exception("CreateDesktop loi " + Marshal.GetLastWin32Error());
    STARTUPINFO si = new STARTUPINFO(); si.cb = Marshal.SizeOf(typeof(STARTUPINFO)); si.lpDesktop = "WinSta0\\" + desk;
    PROCESS_INFORMATION pi;
    if (!CreateProcess(null, new StringBuilder(cmdline), IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero, null, ref si, out pi)) { int e = Marshal.GetLastWin32Error(); CloseDesktop(d); throw new Exception("CreateProcess loi " + e); }
    CloseHandle(pi.hThread); CloseHandle(pi.hProcess);
    System.Threading.Thread.Sleep(1500); // để tiến trình kịp gắn vào desktop trước khi đóng tay cầm
    CloseDesktop(d);
    return pi.dwProcessId;
  }
  public static List<string> Windows(string desk) {
    var r = new List<string>(); IntPtr d = OpenDesktop(desk, 0, false, ALL);
    if (d == IntPtr.Zero) { r.Add("(khong co desktop " + desk + ")"); return r; }
    EnumDesktopWindows(d, (h, l) => { if (IsWindowVisible(h)) { var t = new StringBuilder(256); var c = new StringBuilder(256); GetWindowText(h, t, 256); GetClassName(h, c, 256); uint pid; GetWindowThreadProcessId(h, out pid); r.Add(pid + " | " + c + " | " + t); } return true; }, IntPtr.Zero);
    CloseDesktop(d); return r;
  }
}
'@
if ($LietKe) { [DesktopRieng]::Windows($Ten); exit 0 }
if (-not $Lenh) { Write-Output "Thieu -Lenh"; exit 1 }
$cmd = if ($ThamSo) { "`"$Lenh`" $ThamSo" } else { "`"$Lenh`"" }
$procId = [DesktopRieng]::Run($Ten, $cmd)
Write-Output "PID $procId tren desktop $Ten"
