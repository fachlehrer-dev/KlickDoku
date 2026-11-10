# Drittanbieter-Komponenten

KlickDoku selbst steht unter der MIT-Lizenz (siehe `LICENSE`). Es nutzt die
folgenden Bibliotheken (Lizenzangaben laut PyPI-Metadaten, Stand Nov. 2026 –
bitte vor einer Veröffentlichung von Binärdateien gegen die aktuelle Version prüfen):

| Paket | Lizenz | Plattform |
|---|---|---|
| customtkinter | MIT | beide |
| mss | MIT | beide |
| Pillow | MIT-CMU (HPND) | beide |
| reportlab | BSD | beide |
| pynput | **LGPL-3.0** | beide |
| uiautomation | Apache-2.0 | Windows |
| pyobjc-framework-Quartz / -ApplicationServices | MIT | macOS |
| PyInstaller (nur Build-Werkzeug) | GPL-2.0-or-later mit Ausnahme für erzeugte Programme | beide |

## Hinweise

* **pynput (LGPL-3.0):** Wird unverändert als eigene Bibliothek verwendet.
  Der vollständige Quellcode von KlickDoku liegt offen vor und lässt sich mit
  `build.bat` / `build.sh` jederzeit neu bauen, wodurch pynput austauschbar bleibt.
  Wer fertige EXE/DMG-Dateien weitergibt, legt den LGPL-3.0-Text (https://www.gnu.org/licenses/lgpl-3.0.html)
  und einen Hinweis auf dieses Repository bei.
* **PyInstaller:** Die Ausnahme in der PyInstaller-Lizenz erlaubt es, die damit
  gebauten Programme unter beliebiger Lizenz zu verbreiten.
* **Schriften im PDF:** Der PDF-Export bettet (als Teilmenge) eine auf dem
  jeweiligen Rechner vorhandene Schrift ein (Segoe UI, Arial oder DejaVu Sans).
  KlickDoku liefert selbst keine Schriftdateien mit.
* Dies ist eine technische Zusammenstellung, keine Rechtsberatung.
