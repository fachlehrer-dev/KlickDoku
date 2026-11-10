@echo off
rem ============================================================
rem  Laedt alle Pakete fuer den Offline-Build nach wheels\
rem  Auf einem Rechner MIT Internet ausfuehren, dann den Ordner wheels\
rem  zusammen mit dem Rest auf den Ziel-Rechner kopieren.
rem  WICHTIG: gleiche Windows-Version (64 Bit) und gleiche Python-Version
rem  (z. B. 3.12) wie auf dem Ziel-Rechner verwenden.
rem ============================================================
setlocal
cd /d "%~dp0"
set "PY=python"
where py >nul 2>&1 && set "PY=py -3"
%PY% --version || goto :fail
%PY% -m pip download -d wheels pyinstaller mss pillow pynput customtkinter reportlab uiautomation || goto :fail
echo.
echo Fertig. Ordner wheels\ auf den Ziel-Rechner kopieren und dort build.bat starten.
pause
exit /b 0
:fail
echo [FEHLER] Download fehlgeschlagen.
pause
exit /b 1
