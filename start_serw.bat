@echo off
REM Ścieżka do katalogu z plikami HTML
set CONTENT_DIR=C:\lech_dane\python\wszystkie\output

REM Przejdź do katalogu zawartości
cd /d "%CONTENT_DIR%"

REM Uruchom lokalny serwer HTTP
python -m http.server 8000

REM Zatrzymaj konsolę po zamknięciu serwera
pause