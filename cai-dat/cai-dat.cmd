@echo off
chcp 65001 >nul
title Cai dat AI Da Nguon
set "DIR=%LOCALAPPDATA%\AiDaNguon"
if not exist "%DIR%" mkdir "%DIR%"
echo Dang tai manifest tu GitHub Pages...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -UseBasicParsing 'https://dunghoang1985.github.io/ai-da-nguon/manifest.xml' -OutFile '%DIR%\manifest.xml'; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
  echo.
  echo [LOI] Khong tai duoc manifest. Kiem tra GitHub Pages da bat chua:
  echo https://github.com/DungHoang1985/ai-da-nguon/settings/pages
  pause
  exit /b 1
)
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AiDaNguon /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
if errorlevel 1 (
  echo [LOI] Khong ghi duoc dang ky vao Windows.
  pause
  exit /b 1
)
echo.
echo DA CAI XONG. Dong han Word, Excel, PowerPoint roi mo lai.
echo Vao the Home, bam nut "Mo AI Da Nguon".
pause
