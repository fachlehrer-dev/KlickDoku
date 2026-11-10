#!/usr/bin/env bash
# Laedt alle Pakete fuer den Offline-Build nach wheels/
# Auf einem Mac MIT Internet ausfuehren (gleiche Python-Version und
# Prozessor-Architektur wie auf dem Ziel-Mac), dann wheels/ dorthin kopieren.
set -euo pipefail
cd "$(dirname "$0")"
PY="${PYTHON:-python3}"
"$PY" -m pip download -d wheels pyinstaller mss pillow pynput customtkinter reportlab \
  pyobjc-framework-Quartz pyobjc-framework-ApplicationServices
echo "Fertig. Ordner wheels/ auf den Ziel-Mac kopieren und dort ./build.sh starten."
