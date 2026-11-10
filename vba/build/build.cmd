@echo off
rem Nhay dup de build add-in VBA "Nhat ky TC" va chay kiem thu (can Excel tren Windows).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1" %*
echo.
echo Ket qua da ghi vao: %~dp0..\dist\ket-qua-build.txt
pause
