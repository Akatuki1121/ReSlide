@echo off
chcp 65001 >nul
"%~dp0ReSlide.exe" --gui %*
pause
