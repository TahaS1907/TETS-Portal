@echo off
setlocal
cd /d "%~dp0"
echo [1/3] Sanal ortam olusturuluyor...
where py >nul 2>nul && (py -3 -m venv .venv) || (python -m venv .venv)
if not exist ".venv\Scripts\python.exe" (
  echo Python bulunamadi. python.org adresinden Python 3.10+ kurun ^("Add to PATH" isaretli^).
  pause
  exit /b 1
)
echo [2/3] Paketler kuruluyor...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo Paket kurulumu basarisiz. Sirket proxy ayarini kontrol edin: set HTTPS_PROXY=http://proxy:port
  pause
  exit /b 1
)
echo [3/3] Tarayici kontrolu: config.yaml varsayilani Microsoft Edge (ek indirme gerekmez).
echo Edge yoksa su komutla Chromium indirilebilir: .venv\Scripts\python.exe -m playwright install chromium
if not exist config.yaml copy config.ornek.yaml config.yaml >nul
echo.
echo Kurulum tamam. config.yaml icindeki qlik ayarlarini doldurun, sonra BASLAT.bat ile acin.
pause
