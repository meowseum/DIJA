@echo off
setlocal
cd /d %~dp0

py -m PyInstaller DIJA.spec

echo.
echo Build finished. Check the dist folder.
pause
