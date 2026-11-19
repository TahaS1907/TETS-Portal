@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto missing
.venv\Scripts\python.exe portal_giris.py
if errorlevel 1 pause
exit /b
:missing
echo Once KURULUM.bat dosyasini calistirin.
pause
exit /b 1
