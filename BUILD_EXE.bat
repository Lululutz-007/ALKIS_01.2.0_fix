@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3.13 -m venv .venv
if errorlevel 1 goto error
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
if errorlevel 1 goto error
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
python -m PyInstaller --noconfirm --clean ALKIS_GeoAS_GEOgraf.spec
if errorlevel 1 goto error
powershell -NoProfile -Command "Get-FileHash '.\dist\ALKIS_GeoAS_GEOgraf_V1.2.0.exe' -Algorithm SHA256 | Format-List | Out-File -Encoding utf8 '.\dist\SHA256.txt'"
echo BUILD ERFOLGREICH: dist\ALKIS_GeoAS_GEOgraf_V1.2.0.exe
pause
exit /b 0
:error
echo BUILD FEHLGESCHLAGEN
pause
exit /b 1
