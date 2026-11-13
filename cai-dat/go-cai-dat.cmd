@echo off
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v c95713b4-2c26-42c9-944c-cdbd5f4ff7e9 /f
del /q "%LOCALAPPDATA%\AiDaNguon\manifest.xml" 2>nul
echo Da go AI Da Nguon. Dong va mo lai Office.
pause
