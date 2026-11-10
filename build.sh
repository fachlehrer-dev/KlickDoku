#!/usr/bin/env bash
# ============================================================
#  KlickDoku - baut dist/KlickDoku.app und dist/KlickDoku.dmg (macOS)
#  Aufruf:  ./build.sh        (einmalig:  chmod +x build.sh)
#
#  Die .app ist ein Ordner-Bundle: sie wird beim Start NICHT entpackt und
#  startet deshalb am schnellsten. Das DMG ist stark komprimiert (LZMA).
#  Benoetigt Python 3.9+ mit Tk (python.org-Installer oder: brew install python-tk).
#  Alle Pakete werden automatisch installiert (kein requirements.txt noetig).
# ============================================================
set -euo pipefail
cd "$(dirname "$0")"

if [[ "$(uname)" != "Darwin" ]]; then
  echo "[FEHLER] Dieses Skript baut die macOS-App und muss auf einem Mac laufen." >&2
  exit 1
fi
[[ -f klickdoku.py ]] || { echo "[FEHLER] klickdoku.py liegt nicht neben build.sh." >&2; exit 1; }

APP="KlickDoku"
BUNDLE_ID="${BUNDLE_ID:-com.klickdoku.app}"
PY="${PYTHON:-python3}"

command -v "$PY" >/dev/null || { echo "[FEHLER] $PY nicht gefunden." >&2; exit 1; }
"$PY" -c "import tkinter" 2>/dev/null || {
  echo "[FEHLER] Python hat kein Tk. Python von python.org installieren oder: brew install python-tk" >&2
  exit 1
}

echo "[1/5] Virtuelle Umgebung ..."
[[ -d .venv ]] || "$PY" -m venv .venv
# shellcheck disable=SC1091
source .venv/bin/activate

echo "[2/5] Pakete installieren ..."
PKGS=(pyinstaller mss pillow pynput customtkinter reportlab
      pyobjc-framework-Quartz pyobjc-framework-ApplicationServices)
if [[ -d wheels ]]; then
  echo "    Offline-Modus: verwende den Ordner wheels/"
  python -m pip install --no-index --find-links wheels "${PKGS[@]}"
else
  python -c "import socket; socket.gethostbyname('pypi.org')" 2>/dev/null || {
    echo "[WARNUNG] pypi.org wird nicht aufgeloest (kein Internet, DNS oder Proxy). Versuche es trotzdem ..." >&2
    echo "          Falls es scheitert: auf einem Mac MIT Internet ./fetch_wheels.sh ausfuehren," >&2
    echo "          den Ordner wheels/ neben build.sh kopieren und erneut starten." >&2
  }
  python -m pip install --retries 1 --timeout 10 "${PKGS[@]}"
fi

echo "[3/5] Icon erzeugen ..."
python make_icon.py

# Alles Unnoetige weglassen: kleiner = schnellerer Start
EXCL=()
for m in numpy pandas matplotlib scipy IPython pytest setuptools pkg_resources \
         distutils unittest pydoc doctest lib2to3 xmlrpc sqlite3 ssl _ssl _hashlib \
         lxml PIL.AvifImagePlugin PIL._avif PIL.WebPImagePlugin PIL._webp; do
  EXCL+=(--exclude-module "$m")
done

echo "[4/5] App bauen ..."
rm -rf "dist/$APP.app" build
python -m PyInstaller --noconfirm --clean --windowed --noupx --optimize 2 \
  --name "$APP" \
  --icon assets/icon.icns \
  --osx-bundle-identifier "$BUNDLE_ID" \
  --add-data "assets/icon.png:assets" \
  --add-data "assets/icon_256.png:assets" \
  --add-data "LIZENZHINWEISE.txt:." \
  --add-data "licenses:licenses" \
  --collect-all customtkinter \
  --hidden-import pynput.keyboard._darwin \
  --hidden-import pynput.mouse._darwin \
  "${EXCL[@]}" \
  klickdoku.py

# Ad-hoc-Signatur: macOS behaelt damit die erteilten Berechtigungen eher bei
codesign --force --deep --sign - "dist/$APP.app" || echo "(Signieren uebersprungen)"

echo "[5/5] DMG erstellen ..."
STAGE="dist/dmg"
rm -rf "$STAGE" "dist/$APP.dmg"
mkdir -p "$STAGE"
cp -R "dist/$APP.app" "$STAGE/"
cp LIZENZHINWEISE.txt "$STAGE/"
cp -R licenses "$STAGE/licenses"
ln -s /Applications "$STAGE/Applications"
hdiutil create -volname "$APP" -srcfolder "$STAGE" -ov -format ULMO "dist/$APP.dmg" \
  || hdiutil create -volname "$APP" -srcfolder "$STAGE" -ov -format UDZO "dist/$APP.dmg"
rm -rf "$STAGE"

echo
echo "Fertig:"
echo "  App: $(pwd)/dist/$APP.app   ($(du -sh "dist/$APP.app" | cut -f1))"
echo "  DMG: $(pwd)/dist/$APP.dmg   ($(du -sh "dist/$APP.dmg" | cut -f1))"
echo
echo "Beim ersten Start unter Systemeinstellungen > Datenschutz & Sicherheit"
echo "Bedienungshilfen, Eingabeueberwachung und Bildschirmaufnahme erlauben."
echo "Falls macOS die App blockiert: Rechtsklick > Oeffnen."
