@echo off
chcp 65001 >nul
"%~dp0ReSlide\ReSlide.exe" --gui %*
pause
