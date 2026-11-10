# KlickDoku

Schrittaufzeichnung für **Windows und macOS**: KlickDoku zeichnet deine Klicks mit Screenshot auf (roter Kreis an der Klickstelle, grüner Rahmen um das angeklickte Element) und erzeugt daraus automatisch eine Anleitung als HTML, Word und PDF. Eine Alternative zum Windows-Problemaufzeichnungsprogramm – für beide Systeme, mit Bildeditor und Export für Wikis und Ticketsysteme.

## Funktionen

- Aufnahme von Klicks, Tastenkürzeln und eigenen Hinweisen, Zwischenüberschriften, Pause
- Aufnahmebereich wählbar: Bildschirm des Klicks, alle Bildschirme, ein bestimmter Bildschirm, Fenster des Klicks, ein bestimmtes Fenster aus der Liste, fester Bereich. Klicks außerhalb werden **nicht** aufgezeichnet.
- Prüf-Dialog: Texte ändern, Schritte verschieben, löschen oder an beliebiger Stelle einfügen (auch Schritte ohne Klick)
- Bildeditor: Textmarker, Rahmen, Pfeil (mit Farbwahl), Verwischen
- Export in einem Rutsch in alle gewählten Formate, automatische Nummerierung
- Keine Koordinaten, Zeitstempel, Benutzer- oder Computernamen im Bericht

## Ausgabe

Nach „Beenden“ wählst du die Formate und einen Ordner. Darin entsteht ein Unterordner mit dem Berichtstitel:

| Datei | Inhalt |
|---|---|
| `Titel.html` | formatierte Seite (Dark Mode, Inhaltsverzeichnis, Bilder zoombar) |
| `Titel_einfach.html` | **Standard-HTML** ohne CSS/JS (`h1`, `h2`, `ol/li`, `img`) zum Kopieren in Ticket- oder FAQ-Systeme |
| `Titel_zum_Einfuegen.html` | wie die Standard-HTML, aber **ohne Titel** und mit **eingebetteten Bildern** (Base64). Im Browser öffnen, Strg+A, Strg+C und in einen Editor mit automatischem Bild-Upload (z. B. TinyMCE) einfügen |
| `img_Titel/schritt_01.png …` | die Screenshots (PNG), von den HTML-Dateien per relativer URL eingebunden |
| `Titel.docx` | Word, mit echten Überschriften |
| `Titel.pdf` | PDF |

Der Ordner `img_Titel` muss beim Weitergeben der HTML-Dateien daneben liegen (oder unter „Optionen“ die Bilder einbetten). Word und PDF enthalten die Bilder immer selbst.

**Nummerierung läuft automatisch:** im formatierten HTML per CSS, in Word als echte Nummerierung, im einfachen HTML als `<ol>`. Löschst du später einen Schritt, zählt alles von selbst neu. Nach jeder Zwischenüberschrift beginnt die Nummerierung standardmäßig neu (abschaltbar).

## Bedienung

- **Berichtstitel:** Das Feld oben wird die Überschrift (H1) und der Ordnername.
- **Abschnitt:** Zwischenüberschrift einfügen (Knopf oder Hotkey).
- **Hinweis:** Freitext ohne Schrittnummer.
- **Prüf-Dialog beim Beenden:** Formate wählen; Texte ändern, ↑↓ verschieben, ✕ löschen, ¶ Abschnitt davor, **＋** Schritt davor einfügen („＋ Schritt am Ende“ unten). Vorschaubild anklicken öffnet den Bildeditor.
- **Bildeditor:** Textmarker und Rahmen (Rechteck aufziehen), Pfeil (ziehen), Verwischen (Bereich aufziehen, wird verpixelt). Farbwahl, ↶ Rückgängig (Strg+Z), „Übernehmen“.
- Oben rechts öffnet **?** ein Info-Fenster (Version, Herausgeber, Autor, Link zum Quellcode); ◐ wechselt zwischen hell und dunkel. Die Versionsnummer steht auch in der Titelleiste.
- Einstellungen werden gespeichert.

| | Windows | macOS |
|---|---|---|
| Abschnitt einfügen | Strg+Umschalt+F8 | Ctrl+Shift+H |
| Pause / Fortsetzen | Strg+Umschalt+F9 | Ctrl+Shift+P |
| Hinweis | Strg+Umschalt+F10 | Ctrl+Shift+C |
| Beenden | Strg+Umschalt+F11 | Ctrl+Shift+S |

Windows: Das Programmfenster ist selbst nicht in den Screenshots zu sehen. macOS: Das Fenster wird beim Start minimiert; beenden per Hotkey.

## Aufnahmebereich im Detail

- **Bildschirm des Klicks** (Standard) – Screenshot des Monitors, auf dem geklickt wurde
- **Alle Bildschirme** – ein breites Bild
- **Nur Bildschirm 1, 2, 3 …** – Klicks auf anderen Monitoren werden ignoriert. Die Auswahlliste zeigt die Auflösung; beim Auswählen (und mit dem 🖥-Knopf daneben) wird auf jedem Monitor kurz seine Nummer eingeblendet, der gewählte ist farbig hervorgehoben
- **Fenster des Klicks** – Screenshot nur vom angeklickten Fenster
- **Nur bestimmtes Fenster …** – Fenster aus der Liste offener Fenster wählen (↻ aktualisiert). Es wird beim Start in den Vordergrund geholt; Klicks und Eingaben in anderen Anwendungen werden ignoriert, Menüs und Dialoge derselben Anwendung zählen mit.
- **Nur fester Bereich …** – Rechteck aufziehen; Klicks außerhalb werden ignoriert.

## Installation und Start (zum Entwickeln)

Python 3.9 oder neuer.

```
pip install -r requirements.txt
python klickdoku.py
```

Optionale Pakete liefern Element- und Fensternamen („Linksklick auf ‚Speichern‘ (Button)“): Windows `uiautomation`, macOS `pyobjc-…`. Ohne sie steht nur „Linksklick in ‚Fenster‘“.

## macOS: Berechtigungen

Systemeinstellungen → Datenschutz & Sicherheit: Das Programm, aus dem du startest (Terminal oder die gebaute App), braucht **Bedienungshilfen** (Elementnamen), **Eingabeüberwachung** (Klicks, Tasten) und **Bildschirmaufnahme** (Screenshots, Fenstertitel). Danach neu starten.

## Datenschutz

Getippter Text wird **nicht** aufgezeichnet (Passwörter!), nur „Text eingegeben“ und Sondertasten/Kürzel wie Strg+S oder ENTER. Screenshots zeigen aber alles Sichtbare: Schritte im Prüf-Dialog löschen oder mit dem Bildeditor verwischen, bevor du Anleitungen weitergibst. Es werden keine Daten übertragen; alles bleibt lokal.

## Als EXE / DMG bauen

Den **ganzen Ordner** entpackt lassen (Skripte und `assets/` müssen zusammen liegen). Die Skripte installieren alle Pakete selbst.

- **Windows:** `build.bat` doppelklicken → `dist\KlickDoku.exe` (eine Datei). `build.bat onedir` baut stattdessen einen Ordner ohne Entpacken beim Start (schnellster Start).
- **macOS:** `chmod +x build.sh && ./build.sh` → `dist/KlickDoku.app` und `dist/KlickDoku.dmg`.
- **GitHub Actions:** `.github/workflows/build.yml` baut EXE und DMG per Hand-Start oder bei einem Tag `v*` und legt sie als Artefakte ab.

Ohne Internet auf dem Build-Rechner: auf einem Rechner mit Internet `fetch_wheels.bat` / `./fetch_wheels.sh` ausführen und den Ordner `wheels/` neben die Build-Skripte kopieren (Windows- und Python-Version müssen übereinstimmen). Hinter einem Proxy vorher `set HTTPS_PROXY=http://proxy:8080` setzen; blockiert ein DNS-Filter `pypi.org`, dort `pypi.org` und `files.pythonhosted.org` freigeben.

Die Mac-App ist nur ad-hoc signiert: Beim ersten Start ggf. Rechtsklick → Öffnen und die Berechtigungen der App erteilen. Der Mac-Build gilt für die Architektur des Build-Rechners.

Icon ändern: `make_icon.py` anpassen oder eigene Dateien nach `assets/` legen und die Zeile mit `make_icon.py` in den Build-Skripten entfernen.

## Weitergeben (Website, Download)

Wer `KlickDoku.exe` oder `KlickDoku.dmg` anbietet, legt **`LIZENZHINWEISE.txt`** und den Ordner **`licenses/`** dazu (die Build-Skripte kopieren beides neben die EXE bzw. in die DMG) und trägt in `LIZENZHINWEISE.txt` die Adresse des Quellcodes ein. Hintergrund: Die Bibliothek pynput steht unter LGPL-3.0.

## Lizenz

[MIT](LICENSE). Drittanbieter-Bibliotheken und ihre Lizenzen (u. a. pynput unter LGPL-3.0): siehe [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md).
