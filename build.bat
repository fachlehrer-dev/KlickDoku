@echo off
rem ============================================================
rem  KlickDoku - baut dist\KlickDoku.exe  (Windows)
rem
rem    build.bat           -> eine einzelne, schlanke EXE (Standard)
rem    build.bat onedir    -> Ordner dist\KlickDoku\ (schnellster Start,
rem                           kein Entpacken beim Starten)
rem
rem  Doppelklick genuegt. Benoetigt Python 3.9+ (python.org) mit Tcl/Tk.
rem  Alle Pakete werden automatisch installiert (kein requirements.txt noetig).
rem  Ohne Internet: erst fetch_wheels.bat auf einem Rechner mit Internet ausfuehren
rem  und den Ordner wheels\ neben dieses Skript kopieren.
rem ============================================================
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "MODE=--onefile"
if /i "%~1"=="onedir" set "MODE=--onedir"

set "PY=python"
where py >nul 2>&1 && set "PY=py -3"
%PY% --version >nul 2>&1 || (
  echo [FEHLER] Python nicht gefunden. Bitte von https://www.python.org installieren
  echo          und beim Installieren "Add python.exe to PATH" ankreuzen.
  goto :fail
)
if not exist "klickdoku.py" (
  echo [FEHLER] klickdoku.py liegt nicht neben build.bat.
  goto :fail
)

echo [1/4] Virtuelle Umgebung ...
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv || goto :fail
)
call ".venv\Scripts\activate.bat" || goto :fail

echo [2/4] Pakete installieren ...
set "PKGS=pyinstaller mss pillow pynput customtkinter reportlab uiautomation"
if exist "wheels\" (
  echo     Offline-Modus: verwende den Ordner wheels\
  python -m pip install --no-index --find-links wheels %PKGS% || goto :fail
) else (
  python -c "import socket; socket.gethostbyname('pypi.org')" 2>nul || (
    echo.
    echo [WARNUNG] pypi.org wird nicht aufgeloest ^(kein Internet, DNS oder Proxy^). Versuche es trotzdem ...
    echo          Falls es scheitert: Auf einem Rechner MIT Internet fetch_wheels.bat ausfuehren,
    echo          den entstandenen Ordner wheels\ neben build.bat kopieren und erneut starten.
    echo.
  )
  python -m pip install --retries 1 --timeout 10 %PKGS% || goto :fail
)

echo [3/4] Icon erzeugen ...
python make_icon.py || goto :fail

rem Alles Unnoetige weglassen: kleiner = schnelleres Entpacken/Starten
set "EXCL=--exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module pytest --exclude-module setuptools --exclude-module pkg_resources --exclude-module distutils --exclude-module unittest --exclude-module pydoc --exclude-module doctest --exclude-module lib2to3 --exclude-module xmlrpc --exclude-module sqlite3 --exclude-module ssl --exclude-module _ssl --exclude-module _hashlib --exclude-module lxml --exclude-module PIL.AvifImagePlugin --exclude-module PIL._avif --exclude-module PIL.WebPImagePlugin --exclude-module PIL._webp"

echo [4/4] EXE bauen (%MODE%) ...
python -m PyInstaller --noconfirm --clean %MODE% --windowed --noupx --optimize 2 ^
  --name KlickDoku ^
  --icon "assets\icon.ico" ^
  --add-data "assets\icon.ico;assets" ^
  --add-data "assets\icon_256.png;assets" ^
  --add-data "LIZENZHINWEISE.txt;." ^
  --add-data "licenses;licenses" ^
  --collect-all customtkinter ^
  --collect-all uiautomation ^
  --hidden-import pynput.keyboard._win32 ^
  --hidden-import pynput.mouse._win32 ^
  %EXCL% ^
  klickdoku.py || goto :fail

copy /Y LIZENZHINWEISE.txt dist\ >nul
if exist "dist\licenses" rmdir /S /Q "dist\licenses"
xcopy /E /I /Y /Q licenses dist\licenses >nul
echo.
if /i "%MODE%"=="--onedir" (
  echo Fertig:  %~dp0dist\KlickDoku\KlickDoku.exe
) else (
  echo Fertig:  %~dp0dist\KlickDoku.exe
  for %%F in ("dist\KlickDoku.exe") do echo Groesse: %%~zF Bytes
)
echo.
pause
exit /b 0

:fail
echo.
echo [FEHLER] Der Build ist fehlgeschlagen (siehe Meldungen oben).
pause
exit /b 1
