@echo off
rem Hadislar: Word uchun qo'shimchani o'rnatish (administrator huquqi kerak emas).
rem cmd'da xuddi shu: powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://hadis.abdulloh.org/win.txt | iex"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072; try { $s = irm https://hadis.abdulloh.org/win.txt } catch { $s = irm https://abdullohukr.github.io/hadis/win.txt }; iex $s"
set rc=%errorlevel%
echo.
pause
exit /b %rc%
