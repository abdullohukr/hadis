@echo off
rem Hadislar: tuzatish (toza o'rnatish): Word'da qo'shimcha ishlamasa yoki qotib qolsa (administrator huquqi kerak emas).
rem cmd'da xuddi shu: powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://hadis.abdulloh.org/win-clean.txt | iex"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072; try { $s = irm https://hadis.abdulloh.org/win-clean.txt } catch { $s = irm https://abdullohukr.github.io/hadis/win-clean.txt }; iex $s"
set rc=%errorlevel%
echo.
pause
exit /b %rc%
