@echo off
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v AiDaNguon /f
del /q "%LOCALAPPDATA%\AiDaNguon\manifest.xml" 2>nul
echo Da go AI Da Nguon. Dong va mo lai Office.
pause
