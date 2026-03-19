@echo off
cd /d %~dp0
set TAURI_SIGNING_PRIVATE_KEY_PATH=C:\Users\Jeff\Desktop\Project_S\DIJ\~\.tauri\dija.key
cargo tauri build
echo.
echo Build complete. Output: src-tauri\target\release\bundle\
pause
