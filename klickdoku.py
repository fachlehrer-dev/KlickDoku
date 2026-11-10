#!/usr/bin/env python3
"""
KlickDoku - Schrittaufzeichnung zu Anleitungen (Alternative zum Windows-
Problemaufzeichnungsprogramm).
Laeuft auf Windows und macOS.

Zeichnet Mausklicks (mit Screenshot, rot markierter Klickstelle und, wenn
moeglich, dem Namen des angeklickten Elements) sowie Tastenkuerzel auf und
erzeugt einen Bericht als HTML, Word (.docx) oder PDF.

Datenschutz: Getippter Text wird NICHT aufgezeichnet, nur der Hinweis
"Text eingegeben" sowie Sondertasten/Kuerzel (z. B. Strg+S, Enter).

Hotkeys (global), Windows:  Strg+Umschalt+F9 / F10 / F11
                            = Pause / Kommentar / Beenden & speichern
Hotkeys (global), macOS:    Ctrl+Shift+P / C / S
                            = Pause / Kommentar / Beenden & speichern

Start:  python klickdoku.py [--lang de|en]
"""
from __future__ import annotations

import argparse
import base64
import contextlib
import datetime as dt
import html
import io
import locale
import os
import queue
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass

IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
MAX_W = 1600      # Screenshots werden auf diese Breite verkleinert

# --------------------------------------------------------------------------
# Texte
# --------------------------------------------------------------------------
TEXTS = {
    "de": dict(
        app="KlickDoku",
        start="Aufzeichnung starten", stop="Beenden && speichern",
        pause="Pause", resume="Fortsetzen", comment="Hinweis",
        idle="Bereit", rec="Aufnahme läuft", paused="Pausiert",
        starting="Starte …",
        count="Schritte: {n}",
        hint="Strg+Umschalt+   F8 Abschnitt · F9 Pause\nF10 Hinweis · F11 Ende",
        hint_mac="Ctrl+Shift+   H Abschnitt · P Pause\nC Hinweis · S Ende  "
                 "(Fenster ist beim Aufnehmen minimiert)",
        title_lbl="Berichtstitel",
        title_ph="Titel des Berichts (wird zur Überschrift)",
        heading_btn="Abschnitt", heading_title="Zwischenüberschrift",
        heading_prompt="Titel des neuen Abschnitts:", toc="Inhalt",
        comment_title="Hinweis", comment_prompt="Hinweistext (steht ohne Nummer im Bericht):",
        nothing="Es wurden keine Schritte aufgezeichnet.",
        left="Linksklick", right="Rechtsklick", middle="Mittelklick",
        double="Doppelklick",
        on_elem="{b} auf „{name}“ ({etype})",
        on_pos="{b}",
        in_win=" in „{w}“",
        keys="Tastenkürzel {k}",
        typed="Text eingegeben",
        report_title="Schrittaufzeichnung",
        step="Schritt", created="Stand", comment_lbl="Hinweis",
        no_win="(unbekanntes Fenster)",
        cap="Aufnahmebereich", m_screen="Bildschirm des Klicks",
        m_all="Alle Bildschirme", m_mon="Nur Bildschirm {n}",
        m_window="Fenster des Klicks", m_target="Nur bestimmtes Fenster …",
        m_region="Nur fester Bereich …",
        region_btn="Bereich wählen…", region_info="Bereich: {w}×{h}",
        region_hint="Bereich mit der Maus aufziehen  ·  Esc = Abbrechen",
        pick_window="Bitte zuerst ein Fenster wählen.",
        win_none="(keine Fenster gefunden)", refresh="Fensterliste aktualisieren",
        opts="Ausgabe & Optionen",
        fmt_html="HTML (formatiert)", fmt_plain="HTML (einfach, zum Kopieren)",
        fmt_wiki="HTML zum Einfügen in den Wiki-Editor (Bilder eingebettet)",
        fmt_docx="Word (.docx)", fmt_pdf="PDF",
        opt_restart="Nummerierung je Abschnitt neu beginnen",
        opt_review="Vor dem Speichern prüfen und bearbeiten",
        opt_embed="Bilder in die HTML-Dateien einbetten (statt img_…-Ordner)",
        review_title="Schritte prüfen und bearbeiten",
        review_save="Speichern", review_cancel="Verwerfen",
        review_hint="Texte direkt ändern · ↑↓ verschieben · ✕ löschen · ¶ Abschnitt davor · ＋ Schritt davor · Bild anklicken = bearbeiten",
        add_step="＋ Schritt am Ende", ed_title="Bild bearbeiten", ed_marker="Textmarker",
        ed_frame="Rahmen", ed_arrow="Pfeil", ed_blur="Verwischen", ed_undo="Rückgängig",
        ed_apply="Übernehmen", ed_color="Farbe", ed_hint="Zum Bearbeiten das Bild anklicken", new_step_ph="Neuer Schritt …",
        fmt_choose="Diese Formate erstellen:", export_title="Export",
        note_lbl="Hinweis", section_lbl="Abschnitt",
        choose_dir="Speicherort wählen (es wird ein Unterordner angelegt)",
        discard_q="Ohne Speicherort gehen alle Schritte verloren. Wirklich verwerfen?",
        saved_err="Nicht alle Dateien konnten erstellt werden:\n{e}",
    ),
    "en": dict(
        app="KlickDoku",
        start="Start recording", stop="Stop && save",
        pause="Pause", resume="Resume", comment="Note",
        idle="Ready", rec="Recording", paused="Paused",
        starting="Starting …",
        count="Steps: {n}",
        hint="Ctrl+Shift+   F8 section · F9 pause\nF10 note · F11 stop",
        hint_mac="Ctrl+Shift+   H section · P pause\nC note · S stop  "
                 "(window is minimized while recording)",
        title_lbl="Report title",
        title_ph="Report title (becomes the heading)",
        heading_btn="Section", heading_title="Section heading",
        heading_prompt="Title of the new section:", toc="Contents",
        comment_title="Note", comment_prompt="Note text (appears without a number):",
        nothing="No steps were recorded.",
        left="Left click", right="Right click", middle="Middle click",
        double="Double click",
        on_elem="{b} on “{name}” ({etype})",
        on_pos="{b}",
        in_win=" in “{w}”",
        keys="Keyboard shortcut {k}",
        typed="Text entered",
        report_title="Steps recording",
        step="Step", created="As of", comment_lbl="Note",
        no_win="(unknown window)",
        cap="Capture area", m_screen="Screen of the click",
        m_all="All screens", m_mon="Only screen {n}",
        m_window="Window of the click", m_target="Only a specific window …",
        m_region="Only a fixed region …",
        region_btn="Select region…", region_info="Region: {w}×{h}",
        region_hint="Drag to select a region  ·  Esc = cancel",
        pick_window="Please choose a window first.",
        win_none="(no windows found)", refresh="Refresh window list",
        opts="Output & options",
        fmt_html="HTML (formatted)", fmt_plain="HTML (plain, for copying)",
        fmt_wiki="HTML for pasting into the wiki editor (images embedded)",
        fmt_docx="Word (.docx)", fmt_pdf="PDF",
        opt_restart="Restart numbering in each section",
        opt_review="Review and edit before saving",
        opt_embed="Embed images in the HTML files (instead of img_… folder)",
        review_title="Review and edit steps",
        review_save="Save", review_cancel="Discard",
        review_hint="Edit texts directly · ↑↓ move · ✕ delete · ¶ section before · ＋ insert step before · click image to edit",
        add_step="＋ Step at end", ed_title="Edit image", ed_marker="Highlighter",
        ed_frame="Frame", ed_arrow="Arrow", ed_blur="Blur", ed_undo="Undo",
        ed_apply="Apply", ed_color="Color", ed_hint="Click the image to edit it", new_step_ph="New step …",
        fmt_choose="Create these formats:", export_title="Export",
        note_lbl="Note", section_lbl="Section",
        choose_dir="Choose a location (a subfolder will be created)",
        discard_q="Without a location all steps are lost. Really discard?",
        saved_err="Not all files could be created:\n{e}",
    ),
}


def pick_lang(arg: str | None) -> str:
    if arg in TEXTS:
        return arg
    try:
        loc = (locale.getlocale()[0] or os.environ.get("LANG", "")).lower()
    except Exception:
        loc = ""
    return "de" if loc.startswith(("de", "german")) else "en"


# --------------------------------------------------------------------------
# Datenmodell
# --------------------------------------------------------------------------
@dataclass
class Step:
    kind: str                 # click | key | typed | comment
    ts: dt.datetime
    text: str = ""
    window: str = ""
    image: bytes | None = None
    # nur fuer Klicks:
    button: str = ""
    pos: tuple = (0, 0)
    mono: float = 0.0
    name: str = ""
    etype: str = ""
    double: bool = False


# --------------------------------------------------------------------------
# Windows-Hilfsfunktionen (ctypes)
# --------------------------------------------------------------------------
if IS_WIN:
    from ctypes import (POINTER, Structure, byref, c_int, c_long, c_ulong,
                        c_void_p, c_wchar_p, create_unicode_buffer, windll)
    import ctypes

    class POINT(Structure):
        _fields_ = [("x", c_long), ("y", c_long)]

    _u = windll.user32
    _u.WindowFromPoint.argtypes = [POINT]
    _u.WindowFromPoint.restype = c_void_p
    _u.GetAncestor.argtypes = [c_void_p, ctypes.c_uint]
    _u.GetAncestor.restype = c_void_p
    _u.GetWindowTextLengthW.argtypes = [c_void_p]
    _u.GetWindowTextW.argtypes = [c_void_p, c_wchar_p, c_int]
    _u.GetWindowThreadProcessId.argtypes = [c_void_p, POINTER(c_ulong)]
    _u.GetForegroundWindow.restype = c_void_p
    _u.SetWindowDisplayAffinity.argtypes = [c_void_p, ctypes.c_uint]

    def make_dpi_aware() -> None:
        try:
            windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                _u.SetProcessDPIAware()
            except Exception:
                pass

    def _title(hwnd) -> str:
        if not hwnd:
            return ""
        n = _u.GetWindowTextLengthW(hwnd)
        buf = create_unicode_buffer(n + 1)
        _u.GetWindowTextW(hwnd, buf, n + 1)
        return buf.value

    def _pid(hwnd) -> int:
        pid = c_ulong(0)
        _u.GetWindowThreadProcessId(hwnd, byref(pid))
        return pid.value

    def top_window_at(x: int, y: int):
        h = _u.WindowFromPoint(POINT(int(x), int(y)))
        return _u.GetAncestor(h, 2) if h else None   # GA_ROOT

    def window_title_at(x: int, y: int) -> str:
        return _title(top_window_at(x, y))

    def window_pid_at(x: int, y: int) -> int:
        h = top_window_at(x, y)
        return _pid(h) if h else 0

    def foreground_title() -> str:
        return _title(_u.GetForegroundWindow())

    class RECT(Structure):
        _fields_ = [("left", c_long), ("top", c_long),
                    ("right", c_long), ("bottom", c_long)]

    _u.GetWindowRect.argtypes = [c_void_p, POINTER(RECT)]
    windll.dwmapi.DwmGetWindowAttribute.argtypes = [
        c_void_p, c_ulong, c_void_p, c_ulong]

    def window_rect_at(x: int, y: int):
        """(left, top, breite, hoehe) des Fensters unter (x, y), ohne den
        unsichtbaren Schattenrand."""
        h = top_window_at(x, y)
        if not h:
            return None
        r = RECT()
        if windll.dwmapi.DwmGetWindowAttribute(   # DWMWA_EXTENDED_FRAME_BOUNDS
                h, 9, byref(r), ctypes.sizeof(r)) != 0:
            _u.GetWindowRect(h, byref(r))
        return (r.left, r.top, r.right - r.left, r.bottom - r.top)

    _EnumProc = ctypes.WINFUNCTYPE(c_int, c_void_p, c_void_p)
    _u.EnumWindows.argtypes = [_EnumProc, c_void_p]
    _u.IsWindowVisible.argtypes = [c_void_p]
    _u.IsIconic.argtypes = [c_void_p]
    _u.IsWindow.argtypes = [c_void_p]
    _u.GetWindowLongW.argtypes = [c_void_p, c_int]
    _u.GetWindow.argtypes = [c_void_p, ctypes.c_uint]
    _u.GetWindow.restype = c_void_p
    _u.ShowWindow.argtypes = [c_void_p, c_int]
    _u.SetForegroundWindow.argtypes = [c_void_p]

    def list_windows(own_pid: int):
        """Sichtbare Hauptfenster: [(hwnd, titel, pid)]."""
        out = []

        @_EnumProc
        def cb(hwnd, _lp):
            try:
                if not _u.IsWindowVisible(hwnd) or _u.GetWindow(hwnd, 4):
                    return 1                                   # unsichtbar / Popup
                title = _title(hwnd)
                if not title or title == "Program Manager":
                    return 1
                if _u.GetWindowLongW(hwnd, -20) & 0x80:        # WS_EX_TOOLWINDOW
                    return 1
                cloaked = c_int(0)
                windll.dwmapi.DwmGetWindowAttribute(hwnd, 14, byref(cloaked), 4)
                if cloaked.value:                              # versteckte UWP-Fenster
                    return 1
                pid = _pid(hwnd)
                if pid != own_pid:
                    out.append((int(hwnd), title, pid))
            except Exception:
                pass
            return 1
        _u.EnumWindows(cb, None)
        return out

    def window_rect_of(hwnd):
        if not hwnd or not _u.IsWindow(hwnd) or _u.IsIconic(hwnd):
            return None
        r = RECT()
        if windll.dwmapi.DwmGetWindowAttribute(
                hwnd, 9, byref(r), ctypes.sizeof(r)) != 0:
            _u.GetWindowRect(hwnd, byref(r))
        return (r.left, r.top, r.right - r.left, r.bottom - r.top)

    def foreground_pid() -> int:
        h = _u.GetForegroundWindow()
        return _pid(h) if h else 0

    def activate_window(target) -> None:
        try:
            if _u.IsIconic(target["id"]):
                _u.ShowWindow(target["id"], 9)                 # SW_RESTORE
            _u.SetForegroundWindow(target["id"])
        except Exception:
            pass

    def exclude_from_capture(tk_root) -> None:
        """Eigenes Fenster erscheint nicht in Screenshots (Win10 2004+)."""
        try:
            hwnd = _u.GetAncestor(tk_root.winfo_id(), 2)
            _u.SetWindowDisplayAffinity(hwnd, 0x11)   # WDA_EXCLUDEFROMCAPTURE
        except Exception:
            pass
elif IS_MAC:
    def make_dpi_aware() -> None: ...
    def exclude_from_capture(tk_root) -> None: ...

    def _mac_windows():
        import Quartz
        opts = (Quartz.kCGWindowListOptionOnScreenOnly
                | Quartz.kCGWindowListExcludeDesktopElements)
        return Quartz.CGWindowListCopyWindowInfo(
            opts, Quartz.kCGNullWindowID) or []

    def _mac_window_at(x, y):
        """Vorderstes sichtbares Fenster unter (x, y) (Punkte, oben links 0)."""
        try:
            for w in _mac_windows():          # Reihenfolge: vorne -> hinten
                b = w.get("kCGWindowBounds")
                if not b or w.get("kCGWindowAlpha", 1) < 0.01:
                    continue
                if (b["X"] <= x < b["X"] + b["Width"]
                        and b["Y"] <= y < b["Y"] + b["Height"]):
                    return w
        except Exception:
            pass
        return None

    def _mac_title(w) -> str:
        owner = str(w.get("kCGWindowOwnerName") or "")
        name = str(w.get("kCGWindowName") or "")   # braucht Bildschirmaufnahme-Recht
        return f"{owner} – {name}" if name else owner

    def window_title_at(x: int, y: int) -> str:
        w = _mac_window_at(x, y)
        return _mac_title(w) if w else ""

    def window_pid_at(x: int, y: int) -> int:
        w = _mac_window_at(x, y)
        return int(w.get("kCGWindowOwnerPID", 0)) if w else 0

    def foreground_title() -> str:
        try:
            for w in _mac_windows():
                if w.get("kCGWindowLayer", 1) == 0:
                    return _mac_title(w)
        except Exception:
            pass
        return ""

    def window_rect_at(x: int, y: int):
        w = _mac_window_at(x, y)
        if not w:
            return None
        b = w["kCGWindowBounds"]
        return (int(b["X"]), int(b["Y"]), int(b["Width"]), int(b["Height"]))
    def list_windows(own_pid: int):
        out = []
        try:
            for w in _mac_windows():
                b = w.get("kCGWindowBounds")
                pid = int(w.get("kCGWindowOwnerPID", 0))
                if (w.get("kCGWindowLayer", 1) != 0 or not b or pid == own_pid
                        or b["Width"] < 80 or b["Height"] < 60):
                    continue
                out.append((int(w["kCGWindowNumber"]), _mac_title(w), pid))
        except Exception:
            pass
        return out

    def window_rect_of(wid):
        try:
            for w in _mac_windows():
                if int(w.get("kCGWindowNumber", -1)) == wid:
                    b = w["kCGWindowBounds"]
                    return (int(b["X"]), int(b["Y"]), int(b["Width"]),
                            int(b["Height"]))
        except Exception:
            pass
        return None

    def foreground_pid() -> int:
        try:
            for w in _mac_windows():
                if w.get("kCGWindowLayer", 1) == 0:
                    return int(w.get("kCGWindowOwnerPID", 0))
        except Exception:
            pass
        return 0

    def activate_window(target) -> None:
        try:
            from AppKit import NSRunningApplication
            app = NSRunningApplication.runningApplicationWithProcessIdentifier_(
                target["pid"])
            if app:
                app.activateWithOptions_(1 << 1)    # IgnoringOtherApps
        except Exception:
            pass
else:
    def make_dpi_aware() -> None: ...
    def window_title_at(x, y) -> str: return ""
    def window_pid_at(x, y) -> int: return 0
    def foreground_title() -> str: return ""
    def window_rect_at(x, y): return None
    def list_windows(own_pid): return []
    def window_rect_of(wid): return None
    def foreground_pid() -> int: return 0
    def activate_window(target) -> None: ...
    def exclude_from_capture(tk_root) -> None: ...


def _describe_element_win(x: int, y: int):
    """Benoetigt das optionale Paket 'uiautomation'."""
    import uiautomation as auto
    c = auto.ControlFromPoint(int(x), int(y))
    if c is None:
        return None
    r = c.BoundingRectangle
    etype = (c.ControlTypeName or "").replace("Control", "")
    return (c.Name or "", etype, (r.left, r.top, r.right, r.bottom))


def _describe_element_mac(x: int, y: int):
    """Accessibility-API (braucht Bedienungshilfen-Recht und pyobjc)."""
    from ApplicationServices import (
        AXUIElementCopyAttributeValue, AXUIElementCopyElementAtPosition,
        AXUIElementCreateSystemWide, AXValueGetValue, kAXValueCGPointType,
        kAXValueCGSizeType)

    def attr(el, name):
        err, val = AXUIElementCopyAttributeValue(el, name, None)
        return val if err == 0 else None

    err, el = AXUIElementCopyElementAtPosition(
        AXUIElementCreateSystemWide(), float(x), float(y), None)
    if err != 0 or el is None:
        return None
    name = ""
    for a in ("AXTitle", "AXDescription", "AXValue", "AXHelp"):
        v = attr(el, a)
        if isinstance(v, str) and v.strip():
            name = v.strip()[:80]
            break
    role = str(attr(el, "AXRole") or "")
    etype = role[2:] if role.startswith("AX") else role
    rect = None
    pos, size = attr(el, "AXPosition"), attr(el, "AXSize")
    if pos is not None and size is not None:
        ok1, pt = AXValueGetValue(pos, kAXValueCGPointType, None)
        ok2, sz = AXValueGetValue(size, kAXValueCGSizeType, None)
        if ok1 and ok2:
            rect = (pt.x, pt.y, pt.x + sz.width, pt.y + sz.height)
    return (name, etype, rect)


def describe_element(x: int, y: int):
    """(Name, Typ, (l,t,r,b)) des UI-Elements unter dem Mauszeiger oder None."""
    try:
        if IS_WIN:
            return _describe_element_win(x, y)
        if IS_MAC:
            return _describe_element_mac(x, y)
    except Exception:
        pass
    return None


def uia_thread_ctx():
    if IS_WIN:
        try:
            import uiautomation as auto
            return auto.UIAutomationInitializerInThread()
        except Exception:
            pass
    return contextlib.nullcontext()


# --------------------------------------------------------------------------
# Rekorder
# --------------------------------------------------------------------------
def resource(*parts: str) -> str:
    """Pfad zu mitgelieferten Dateien (auch in der PyInstaller-App)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def _warm_imports() -> None:
    """Haeufig gebrauchte Module im Hintergrund vorladen (schnellerer 1. Start
    der Aufnahme); Fehler sind egal."""
    try:
        import mss  # noqa: F401
        from PIL import Image, ImageDraw  # noqa: F401
        if IS_WIN:
            from pynput import keyboard, mouse  # noqa: F401
    except Exception:
        pass


def new_mss(mss):
    """mss.MSS ab Version 10, davor mss.mss (dort deprecated/entfernt)."""
    return getattr(mss, "MSS", None)() if hasattr(mss, "MSS") else mss.mss()


def mod_of(key):
    """'ctrl' | 'shift' | 'alt' | 'win' (= Win-/Cmd-Taste) oder None."""
    name = getattr(key, "name", None)
    if not name:
        return None
    base = "alt" if name == "alt_gr" else name.split("_")[0]
    return {"ctrl": "ctrl", "shift": "shift", "alt": "alt",
            "cmd": "win"}.get(base)


HOTKEYS_WIN = {"f8": "heading", "f9": "pause", "f10": "comment",
               "f11": "stop"}
HOTKEYS_MAC = {0x04: "heading", 0x23: "pause", 0x08: "comment",
               0x01: "stop"}   # H, P, C, S
SPECIAL_KEYS = {"enter", "esc", "tab", "delete", "backspace", "print_screen",
                "insert", "home", "end", "page_up", "page_down",
                "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10",
                "f11", "f12"}


class Recorder:
    def __init__(self, lang: str):
        self.T = TEXTS[lang]
        self.steps: list[Step] = []
        self.active = False
        self.paused = False
        self.events: "queue.Queue" = queue.Queue()       # Klick/Taste -> Worker
        self.hotkeys: "queue.Queue[str]" = queue.Queue()  # -> GUI
        self.mods: set[str] = set()
        self.typed_noted = False
        self.own_pid = os.getpid()
        self.mode = "screen"      # screen | all | monitor:N | window | target | region
        self.region = None        # (left, top, breite, hoehe)
        self.target = None        # {"id", "pid", "title"} bei mode == "target"
        self._ml = self._kl = self._worker = None
        self._ready = threading.Event()

    # ---- Steuerung ------------------------------------------------------
    def start(self) -> None:
        from pynput import keyboard, mouse
        self.steps.clear()
        self.mods.clear()
        self.typed_noted = False
        self.active, self.paused = True, False
        self._ready.clear()
        self._worker = threading.Thread(target=self._work, daemon=True)
        self._worker.start()
        # Erst loslegen, wenn Screenshot-/UI-Bibliotheken geladen sind (der
        # Import von uiautomation dauert beim ersten Mal ~1 s) - sonst waere
        # der erste Screenshot zu spaet.
        self._ready.wait(6)
        self._ml = mouse.Listener(on_click=self._on_click)
        self._kl = keyboard.Listener(on_press=self._on_press,
                                     on_release=self._on_release)
        self._ml.start()
        self._kl.start()

    def stop(self) -> None:
        self.active = False
        for lst in (self._ml, self._kl):
            if lst:
                lst.stop()
        self.events.put(None)
        if self._worker:
            self._worker.join(timeout=8)

    def add_comment(self, text: str) -> None:
        self.steps.append(Step("comment", dt.datetime.now(), text=text))

    def add_heading(self, text: str) -> None:
        self.steps.append(Step("heading", dt.datetime.now(), text=text))

    # ---- Listener-Callbacks (duerfen nicht blockieren) ---------------------
    def _on_click(self, x, y, button, pressed):
        if not (pressed and self.active and not self.paused):
            return
        if (IS_WIN or IS_MAC) and window_pid_at(x, y) == self.own_pid:
            return
        self.events.put(("click", x, y, button.name, time.monotonic(),
                         dt.datetime.now()))

    def _mod_label(self, m: str) -> str:
        de = self.T is TEXTS["de"]
        return {"ctrl": "Strg" if de else "Ctrl",
                "alt": "Option" if IS_MAC else "Alt",
                "win": "Cmd" if IS_MAC else "Win",
                "shift": "Umschalt" if de else "Shift"}[m]

    def _hotkey(self, key):
        if not {"ctrl", "shift"} <= self.mods:
            return None
        if IS_MAC:
            return HOTKEYS_MAC.get(getattr(key, "vk", None))
        return HOTKEYS_WIN.get(getattr(key, "name", None))

    def _on_press(self, key):
        mod = mod_of(key)
        if mod:
            self.mods.add(mod)
            return
        act = self._hotkey(key)
        if act:
            self.hotkeys.put(act)
            return
        if not self.active or self.paused:
            return
        name = getattr(key, "name", None)
        now = dt.datetime.now()
        held = set(self.mods) - {"shift"}
        if IS_MAC:
            held.discard("alt")      # Option+Taste erzeugt nur Sonderzeichen
        elif held == {"ctrl", "alt"}:
            held = set()             # AltGr auf deutschen Tastaturen
        if held:
            label = self._key_label(key)
            if label:
                order = [m for m in ("ctrl", "alt", "win", "shift")
                         if m in self.mods]
                combo = "+".join([self._mod_label(m) for m in order] + [label])
                self.events.put(("key", combo, now))
            return
        if name in SPECIAL_KEYS:
            self.events.put(("key", name.upper().replace("_", " "), now))
        elif name in (None, "space") or len(name or "") == 1:
            self.events.put(("typed", now))

    def _on_release(self, key):
        mod = mod_of(key)
        if mod:
            self.mods.discard(mod)

    @staticmethod
    def _key_label(key) -> str:
        name = getattr(key, "name", None)
        if name:
            return name.upper().replace("_", " ")
        ch = getattr(key, "char", None)
        if ch:
            if ord(ch[0]) < 32:                  # Strg+Buchstabe -> Steuerzeichen
                ch = chr(ord(ch[0]) + 64)
            if ch.isprintable() and not ch.isspace():
                return ch.upper()
        vk = getattr(key, "vk", None)
        if IS_WIN and vk and (0x30 <= vk <= 0x39 or 0x41 <= vk <= 0x5A):
            return chr(vk)
        return ""

    # ---- Worker: Screenshot, UI-Element, Schritt anlegen ------------------
    def _work(self) -> None:
        try:
            self._work_loop()
        finally:
            self._ready.set()      # start() darf nie ewig warten

    def _work_loop(self) -> None:
        import mss
        with uia_thread_ctx(), new_mss(mss) as sct:
            self._ready.set()
            while True:
                ev = self.events.get()
                if ev is None:
                    break
                try:
                    if ev[0] == "click":
                        self._handle_click(sct, *ev[1:])
                    elif ev[0] in ("key", "typed"):
                        if (self.mode == "target" and self.target
                                and foreground_pid() != self.target["pid"]):
                            continue          # Tastatur in anderem Fenster
                        win = foreground_title()
                        suffix = (self.T["in_win"].format(w=win) if win else "")
                        if ev[0] == "key":
                            self.steps.append(Step(
                                "key", ev[2], window=win,
                                text=self.T["keys"].format(k=ev[1]) + suffix))
                            self.typed_noted = False
                        elif not self.typed_noted:
                            self.typed_noted = True
                            self.steps.append(Step(
                                "typed", ev[1], window=win,
                                text=self.T["typed"] + suffix))
                except Exception as exc:   # Aufnahme darf nie abstuerzen
                    print("Fehler:", exc, file=sys.stderr)

    def _accepts(self, sct, x, y) -> bool:
        """Klicks ausserhalb des gewaehlten Fensters/Bildschirms/Bereichs
        werden nicht aufgezeichnet."""
        mode = self.mode
        if mode == "region" and self.region:
            l, t, w, h = self.region
            return l <= x < l + w and t <= y < t + h
        if mode.startswith("monitor:"):
            n = int(mode.split(":")[1])
            if 1 <= n < len(sct.monitors):
                m = sct.monitors[n]
                return (m["left"] <= x < m["left"] + m["width"]
                        and m["top"] <= y < m["top"] + m["height"])
        if mode == "target" and self.target:
            return (window_pid_at(x, y) == self.target["pid"]
                    and window_rect_of(self.target["id"]) is not None)
        return True

    def _handle_click(self, sct, x, y, btn, mono, ts) -> None:
        T = self.T
        if not self._accepts(sct, x, y):
            return
        self.typed_noted = False
        last = self.steps[-1] if self.steps else None
        if (last and last.kind == "click" and last.button == btn
                and not last.double and mono - last.mono < 0.5
                and abs(last.pos[0] - x) < 8 and abs(last.pos[1] - y) < 8):
            last.double = True
            last.text = self._click_text(last)
            return
        elem = describe_element(x, y)
        s = Step("click", ts, button=btn, pos=(x, y), mono=mono,
                 window=window_title_at(x, y) or T["no_win"])
        if elem and (elem[0] or elem[1]):
            s.name, s.etype = elem[0], elem[1]
        s.text = self._click_text(s)
        s.image = self._screenshot(sct, x, y, elem[2] if elem else None,
                                   self._capture_rect(sct, x, y))
        self.steps.append(s)

    def _click_text(self, s: Step) -> str:
        T = self.T
        b = T["double"] if s.double else T.get(s.button, s.button)
        if s.name:
            t = T["on_elem"].format(b=b, name=s.name, etype=s.etype)
        else:
            t = T["on_pos"].format(b=b, x=s.pos[0], y=s.pos[1])
        return t + (T["in_win"].format(w=s.window) if s.window else "")

    def _capture_rect(self, sct, x, y):
        """Bereich (left, top, breite, hoehe) in Bildschirmkoordinaten."""
        def of(m):
            return (m["left"], m["top"], m["width"], m["height"])

        def clip(r):
            v = sct.monitors[0]
            l, t = max(r[0], v["left"]), max(r[1], v["top"])
            rr = min(r[0] + r[2], v["left"] + v["width"])
            bb = min(r[1] + r[3], v["top"] + v["height"])
            return (l, t, rr - l, bb - t) if rr - l > 20 and bb - t > 20 else None

        mode = self.mode
        if mode == "all":
            return of(sct.monitors[0])
        if mode.startswith("monitor:"):
            n = int(mode.split(":")[1])
            if 1 <= n < len(sct.monitors):
                return of(sct.monitors[n])
        elif mode == "window":
            r = window_rect_at(x, y)
            c = clip(r) if r else None
            if c:
                return c
        elif mode == "target" and self.target:
            # Zielfenster + ggf. Menue/Dialog derselben Anwendung unter dem Klick
            rects = [r for r in (window_rect_of(self.target["id"]),
                                 window_rect_at(x, y)
                                 if window_pid_at(x, y) == self.target["pid"]
                                 else None) if r]
            if rects:
                l = min(r[0] for r in rects)
                t = min(r[1] for r in rects)
                rr = max(r[0] + r[2] for r in rects)
                bb = max(r[1] + r[3] for r in rects)
                c = clip((l, t, rr - l, bb - t))
                if c:
                    return c
        elif mode == "region" and self.region:
            return self.region
        mons = sct.monitors[1:] or [sct.monitors[0]]
        m = next((m for m in mons
                  if m["left"] <= x < m["left"] + m["width"]
                  and m["top"] <= y < m["top"] + m["height"]), mons[0])
        return of(m)

    @staticmethod
    def _screenshot(sct, x, y, rect, cap) -> bytes:
        from PIL import Image, ImageDraw
        left, top, W, H = cap
        shot = sct.grab({"left": left, "top": top, "width": W, "height": H})
        img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
        f = img.width / W                  # Retina: Pixel pro Punkt (sonst 1)
        W, H = img.size
        limit = MAX_W if W / f <= 2600 else 2800   # mehrere Monitore: etwas groesser
        scale = min(1.0, limit / W)        # Markierungen in Zielgroesse halten
        d = ImageDraw.Draw(img)
        if rect:   # gruener Rahmen um das Element (wenn nicht riesig)
            l, t, r, b = ((rect[0] - left) * f, (rect[1] - top) * f,
                          (rect[2] - left) * f, (rect[3] - top) * f)
            if 0 < (r - l) * (b - t) < 0.5 * W * H and r > 0 and b > 0 \
                    and l < W and t < H:
                d.rectangle((l, t, r, b), outline=(0, 170, 60),
                            width=max(2, int(3 / scale)))
        cx, cy = (x - left) * f, (y - top) * f
        if 0 <= cx < W and 0 <= cy < H:    # Klick ausserhalb: keine Markierung
            rad = int(max(10, 16 / scale))
            d.ellipse((cx - rad, cy - rad, cx + rad, cy + rad),
                      outline=(230, 0, 0), width=max(3, int(4 / scale)))
        if scale < 1:
            img = img.resize((int(W * scale), max(1, int(H * scale))),
                             Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, "PNG", compress_level=4)     # verlustfrei, schnell genug
        return buf.getvalue()


# --------------------------------------------------------------------------
# Berichte: HTML (formatiert), HTML (einfach), Word, PDF
# --------------------------------------------------------------------------
def count_steps(steps) -> int:
    """Nur echte Schritte (keine Abschnitte, keine Hinweise)."""
    return sum(1 for s in steps if s.kind in ("click", "key", "typed", "manual"))


def _stand() -> str:
    return dt.date.today().strftime("%d.%m.%Y")


CSS = """
:root{--bg:#f4f5f7;--card:#fff;--fg:#111827;--mut:#6b7280;--acc:#3b82f6;
--warn:#f59e0b;--bd:#e5e7eb;--note:#fffbeb}
@media(prefers-color-scheme:dark){:root{--bg:#0f1115;--card:#171a21;--fg:#e5e7eb;
--mut:#9ca3af;--bd:#2a2f3a;--note:#2a2210}}
*{box-sizing:border-box}
body{font:15px/1.55 Inter,'Segoe UI',system-ui,-apple-system,sans-serif;margin:0;
background:var(--bg);color:var(--fg)}
main{max-width:980px;margin:0 auto;padding:28px 16px 70px;counter-reset:step}
.hero{padding:26px 28px;border-radius:18px;margin-bottom:22px;color:#fff;
background:linear-gradient(135deg,#2563eb,#7c3aed)}
.hero h1{margin:0 0 10px;font-size:30px;line-height:1.2;letter-spacing:-.02em}
.chips span{background:#ffffff26;border-radius:99px;padding:3px 12px;font-size:13px}
.toc{background:var(--card);border:1px solid var(--bd);border-radius:14px;
padding:14px 22px;margin-bottom:22px}
.toc b{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--mut)}
.toc ol{margin:6px 0 0;padding-left:20px}.toc a{color:var(--acc);text-decoration:none}
.toc a:hover{text-decoration:underline}
h2.sec{font-size:21px;letter-spacing:-.01em;margin:34px 0 12px;padding-bottom:6px;
border-bottom:2px solid var(--acc)}
.step{background:var(--card);border:1px solid var(--bd);border-radius:14px;
padding:14px 18px;margin:0 0 14px;break-inside:avoid;counter-increment:step;
box-shadow:0 1px 2px #0000000d}
.step h3{font-size:16px;font-weight:600;margin:0;display:flex;gap:10px;align-items:baseline}
.n{background:var(--acc);color:#fff;border-radius:99px;padding:1px 10px;font-size:13px;
flex:none;font-weight:600}
.n::before{content:counter(step)}
.note{background:var(--note);border-left:5px solid var(--warn);border-radius:10px;
padding:10px 16px;margin:0 0 14px}
.step img{max-width:100%;border:1px solid var(--bd);border-radius:10px;margin-top:10px;
cursor:zoom-in}
#lb{display:none;position:fixed;inset:0;background:#000d;z-index:9;overflow:auto;
cursor:zoom-out}#lb img{display:block;margin:auto;max-width:none;min-width:100%}
@media print{:root{--bg:#fff;--card:#fff;--fg:#111;--mut:#555;--bd:#ccc}
.hero{background:#2563eb!important;-webkit-print-color-adjust:exact;
print-color-adjust:exact}.step{box-shadow:none}}
"""
CSS_RESTART = "h2.sec{counter-reset:step}\n"   # Nummerierung je Abschnitt neu
JS = ("document.querySelectorAll('.step img').forEach(i=>i.onclick=()=>{"
      "const l=document.getElementById('lb');l.firstChild.src=i.src;"
      "l.style.display='block'});"
      "document.getElementById('lb').onclick=e=>e.currentTarget.style.display='none';")


def make_img_src(img_dir):
    """Liefert img_src(nr, png_bytes) -> URL.
    img_dir = (absoluter Ordner, relativer Ordnername): Bild wird als
    <ordner>/schritt_NN.png gespeichert und per (URL-kodiertem) relativem Pfad
    verlinkt. img_dir = None: Bild wird als Base64 direkt eingebettet."""
    from urllib.parse import quote

    def img_src(nr: int, data: bytes) -> str:
        if img_dir is None:
            return "data:image/png;base64," + base64.b64encode(data).decode("ascii")
        os.makedirs(img_dir[0], exist_ok=True)
        name = f"schritt_{nr:02d}.png"
        with open(os.path.join(img_dir[0], name), "wb") as f:
            f.write(data)
        return quote(f"{img_dir[1]}/{name}")
    return img_src


def build_report(steps: list[Step], title: str, lang: str,
                 restart: bool = True, img_src=None) -> str:
    """Formatiertes HTML. Die Schrittnummern erzeugt der Browser per CSS-Zaehler:
    Loescht man spaeter einen <section class="step"> im Quelltext, zaehlt
    alles automatisch neu."""
    T = TEXTS[lang]
    e = html.escape
    img_src = img_src or make_img_src(None)
    parts, toc = [], []
    nimg = 0
    for s in steps:
        if s.kind == "heading":
            sid = f"sec{len(toc) + 1}"
            toc.append(f'<li><a href="#{sid}">{e(s.text)}</a></li>')
            parts.append(f'<h2 class="sec" id="{sid}">{e(s.text)}</h2>')
        elif s.kind == "comment":
            parts.append(f'<aside class="note"><b>{e(T["comment_lbl"])}:</b> '
                         f'{e(s.text)}</aside>')
        else:
            img = ""
            if s.image:
                nimg += 1
                img = (f'<img alt="{e(s.text)}" '
                       f'src="{e(img_src(nimg, s.image))}">')
            parts.append(f'<section class="step"><h3><span class="n"></span>'
                         f'<span>{e(s.text)}</span></h3>{img}</section>')
    nav = (f'<nav class="toc"><b>{e(T["toc"])}</b><ol>{"".join(toc)}</ol></nav>'
           if toc else "")
    css = CSS + (CSS_RESTART if restart else "")
    return (f'<!doctype html><html lang="{lang}"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1">'
            f"<title>{e(title)}</title><style>{css}</style></head><body><main>"
            f'<header class="hero"><h1>{e(title)}</h1>'
            f'<div class="chips"><span>{e(T["created"])}: {_stand()}</span></div>'
            f'</header>{nav}{"".join(parts)}</main>'
            f'<div id="lb"><img alt=""></div><script>{JS}</script></body></html>')


def write_html(steps, path, title, lang, restart=True, img_dir=None) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(build_report(steps, title, lang, restart, make_img_src(img_dir)))


def build_plain(steps: list[Step], title: str, lang: str, restart: bool,
                img_src, with_title: bool = True) -> str:
    """Schlichtes Standard-HTML ohne CSS/JS: <h1>, <h2>, <ol><li>, <img>.
    Laesst sich in Ticket-/FAQ-Systeme kopieren. img_src(nr, bytes) -> src."""
    T = TEXTS[lang]
    e = html.escape
    out = [f"<h1>{e(title)}</h1>"] if with_title else []
    state = {"ol": False, "li": False, "k": 0, "img": 0}

    def end_li():
        if state["li"]:
            out.append("</li>")
            state["li"] = False

    def end_ol():
        end_li()
        if state["ol"]:
            out.append("</ol>")
            state["ol"] = False

    for s in steps:
        if s.kind == "heading":
            end_ol()
            out.append(f"<h2>{e(s.text)}</h2>")
            if restart:
                state["k"] = 0
        elif s.kind == "comment":
            out.append(f"<p><strong>{e(T['comment_lbl'])}:</strong> {e(s.text)}</p>")
        else:
            end_li()
            if not state["ol"]:
                k = state["k"]
                out.append(f'<ol start="{k + 1}">' if k else "<ol>")
                state["ol"] = True
            state["k"] += 1
            out.append(f"<li><p>{e(s.text)}</p>")
            state["li"] = True
            if s.image:
                state["img"] += 1
                src = img_src(state["img"], s.image)
                out.append(f'<p><img src="{e(src)}" alt="{e(s.text)}"></p>')
    end_ol()
    return (f'<!DOCTYPE html>\n<html lang="{lang}">\n<head>\n<meta charset="utf-8">\n'
            f"<title>{e(title)}</title>\n</head>\n<body>\n" + "\n".join(out)
            + "\n</body>\n</html>\n")


def write_wiki(steps, path, title, lang, restart=True) -> None:
    """Ohne <h1>, Bilder als Base64 -> Seite im Browser oeffnen, Strg+A, Strg+C,
    im Wiki-Editor einfuegen (Editor laedt die Bilder automatisch hoch)."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(build_plain(steps, title, lang, restart, make_img_src(None),
                            with_title=False))


def write_plain(steps, path, title, lang, restart=True, img_dir=None) -> None:
    """img_dir = (absoluter Ordner, relativer Name) -> Bilder als PNG-Dateien;
    None -> Bilder als Base64 direkt im HTML."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(build_plain(steps, title, lang, restart, make_img_src(img_dir)))


# ---- Word ------------------------------------------------------------------
_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def _x(text: str) -> str:
    from xml.sax.saxutils import escape
    return escape(_XML_BAD.sub("", text))


_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def write_docx(steps, path, title, lang, restart=True) -> None:
    """Schlanker DOCX-Export ohne Zusatzpakete (reines OOXML in einer ZIP).
    Die Schrittnummern sind eine echte Word-Nummerierung ("Schritt 1:" ...):
    Loescht man spaeter einen Schritt, zaehlt Word automatisch neu."""
    import zipfile
    from PIL import Image
    T = TEXTS[lang]
    body: list[str] = []
    media: list[bytes] = []

    def run(text, bold=False, color=None, size=None):
        rpr = ("<w:b/>" if bold else "") + (
            f'<w:color w:val="{color}"/>' if color else "") + (
            f'<w:sz w:val="{size}"/>' if size else "")
        rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
        return f'<w:r>{rpr}<w:t xml:space="preserve">{_x(text)}</w:t></w:r>'

    def para(runs, style=None, keep=False, num=None, border=False,
             before=None, after=None, ind=None):
        ppr = (f'<w:pStyle w:val="{style}"/>' if style else "")
        ppr += "<w:keepNext/>" if keep else ""
        ppr += (f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr>'
                if num else "")
        ppr += ('<w:pBdr><w:left w:val="single" w:sz="24" w:space="8" '
                'w:color="F59E0B"/></w:pBdr>' if border else "")
        if before is not None or after is not None:
            ppr += "<w:spacing" + (f' w:before="{before}"' if before is not None
                                   else "") + (f' w:after="{after}"'
                                               if after is not None else "") + "/>"
        ppr += f'<w:ind w:left="{ind}"/>' if ind else ""
        ppr = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
        return f"<w:p>{ppr}{''.join(runs)}</w:p>"

    def picture(data: bytes) -> str:
        w, h = Image.open(io.BytesIO(data)).size
        cx = 16 * 360000                       # 16 cm breit
        cy = int(cx * h / w)
        if cy > 20 * 360000:                   # max. 20 cm hoch
            cx, cy = int(cx * 20 * 360000 / cy), 20 * 360000
        media.append(data)
        n = len(media)
        return (
            '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" '
            f'distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:docPr id="{n}" name="Bild {n}"/><a:graphic><a:graphicData '
            'uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="bild{n}.png"/>'
            f'<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId{n + 2}"/>'
            '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>")

    body.append(para([run(title)], style="Title"))
    body.append(para([run(f"{T['created']}: {_stand()}", color="6B7280", size=18)],
                     after=200))
    num_id = 1                    # jede Gliederungsebene bekommt ggf. eigenen Zaehler
    for idx, s in enumerate(steps):
        if s.kind == "heading":
            body.append(para([run(s.text)], style="Heading1"))
            if restart:
                num_id += 1
        elif s.kind == "comment":
            body.append(para([run(f"{T['comment_lbl']}: ", bold=True, color="D97706"),
                              run(s.text)], border=True, before=60, after=140,
                             ind=220))
        else:
            body.append(para([run(s.text)], keep=bool(s.image), num=num_id,
                             before=240, after=60))
            if s.image:
                body.append(picture(s.image))

    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document xmlns:w="{_W}" xmlns:r="{_R}" '
           'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           f'<w:body>{"".join(body)}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
           'w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body>'
           "</w:document>")
    lang_tag = "de-DE" if lang == "de" else "en-US"
    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{_W}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" '
        'w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
        f'<w:lang w:val="{lang_tag}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        '<w:spacing w:after="80" w:line="264" w:lineRule="auto"/></w:pPr>'
        "</w:pPrDefault></w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
        '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:b/>'
        '<w:color w:val="1D4ED8"/><w:sz w:val="52"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/>'
        '<w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1D4ED8"/>'
        '<w:sz w:val="32"/></w:rPr></w:style></w:styles>')
    numbering = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="{_W}">'
        '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        f'<w:lvlText w:val="{_x(T["step"])} %1:"/><w:suff w:val="space"/>'
        '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>'
        '<w:rPr><w:b/><w:color w:val="2563EB"/></w:rPr></w:lvl></w:abstractNum>'
        + "".join(
            f'<w:num w:numId="{i}"><w:abstractNumId w:val="0"/>'
            + ('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/>'
               "</w:lvlOverride>" if i > 1 else "") + "</w:num>"
            for i in range(1, num_id + 1))
        + "</w:numbering>")
    ctypes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        "</Types>")
    root_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL}"><Relationship Id="rId1" '
        f'Type="{_R}/officeDocument" Target="word/document.xml"/></Relationships>')
    doc_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL}"><Relationship Id="rId1" '
        f'Type="{_R}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId2" Type="{_R}/numbering" Target="numbering.xml"/>'
        + "".join(f'<Relationship Id="rId{n + 2}" Type="{_R}/image" '
                  f'Target="media/bild{n}.png"/>' for n in range(1, len(media) + 1))
        + "</Relationships>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ctypes_xml)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        for n, data in enumerate(media, 1):
            z.writestr(f"word/media/bild{n}.png", data, zipfile.ZIP_STORED)


# ---- PDF -------------------------------------------------------------------
def _pdf_font() -> str:
    """Unicode-Schrift (Windows/macOS/Linux), sonst Helvetica."""
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        win = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        candidates = [
            (os.path.join(win, "segoeui.ttf"), os.path.join(win, "segoeuib.ttf")),
            (os.path.join(win, "arial.ttf"), os.path.join(win, "arialbd.ttf")),
            ("/System/Library/Fonts/Supplemental/Arial.ttf",
             "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
            ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
            ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
             "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ]
        for rp, bp in candidates:
            if os.path.exists(rp) and os.path.exists(bp):
                pdfmetrics.registerFont(TTFont("U", rp))
                pdfmetrics.registerFont(TTFont("U-B", bp))
                pdfmetrics.registerFontFamily("U", normal="U", bold="U-B")
                return "U"
    except Exception:
        pass
    return "Helvetica"


def write_pdf(steps, path, title, lang, restart=True) -> None:
    from PIL import Image as PILImage
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (Image, KeepTogether, Paragraph,
                                    SimpleDocTemplate, Spacer)
    from xml.sax.saxutils import escape as x
    T = TEXTS[lang]
    f = _pdf_font()
    bold = "U-B" if f == "U" else "Helvetica-Bold"
    st_t = ParagraphStyle("t", fontName=bold, fontSize=22, leading=26, spaceAfter=4,
                          textColor="#1d4ed8")
    st_i = ParagraphStyle("i", fontName=f, fontSize=9, leading=12,
                          textColor="#6b7280", spaceAfter=8)
    st_h = ParagraphStyle("h", fontName=f, fontSize=12, leading=15, spaceBefore=14,
                          spaceAfter=4)
    st_s = ParagraphStyle("s", fontName=bold, fontSize=15, leading=19,
                          spaceBefore=20, spaceAfter=4, textColor="#1d4ed8",
                          keepWithNext=1)
    st_n = ParagraphStyle("n", fontName=f, fontSize=11, leading=14, spaceBefore=6,
                          spaceAfter=8, leftIndent=10, borderPadding=4,
                          backColor="#fffbeb")
    doc = SimpleDocTemplate(path, pagesize=A4, title=title,
                            leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=2 * cm, bottomMargin=2 * cm)
    maxw, maxh = 17 * cm, 20 * cm
    story = [Paragraph(x(title), st_t),
             Paragraph(x(f"{T['created']}: {_stand()}"), st_i)]
    n = 0
    for s in steps:
        if s.kind == "heading":
            story.append(Paragraph(x(s.text), st_s))
            if restart:
                n = 0
        elif s.kind == "comment":
            story.append(Paragraph(
                f'<font color="#d97706"><b>{x(T["comment_lbl"])}:</b></font> '
                f"{x(s.text)}", st_n))
        else:
            n += 1
            block = [Paragraph(f'<font color="#2563eb"><b>{x(T["step"])} {n}:'
                               f"</b></font> {x(s.text)}", st_h)]
            if s.image:
                w, h = PILImage.open(io.BytesIO(s.image)).size
                sc = min(maxw / w, maxh / h)
                block.append(Image(io.BytesIO(s.image), w * sc, h * sc))
            story.append(KeepTogether(block))
            story.append(Spacer(1, 4))
    doc.build(story)


# ---- Alles auf einmal exportieren ------------------------------------------
def safe_name(title: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n]+', "_", title).strip(" .")[:80] or "Schritte"


def unique_dir(path: str) -> str:
    cand, i = path, 2
    while os.path.exists(cand):
        cand = f"{path} ({i})"
        i += 1
    return cand


# ---- Bild-Nachbearbeitung (Textmarker, Rahmen, Pfeil, Verwischen) ----------
def _hex_rgb(c: str):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def edit_widths(size):
    """Linienstaerken relativ zur Bildbreite -> (marker, rahmen/pfeil)."""
    w = size[0]
    return max(14.0, w * 0.02), max(3.0, w * 0.003)


def render_edits(img, ops, k: float = 1.0):
    """Wendet ops auf ein PIL-Bild an. Koordinaten/Staerken in Original-Pixeln;
    k skaliert sie (fuer die verkleinerte Vorschau). ops: dict(kind, color,
    pts, w) mit kind marker|frame|arrow|blur."""
    from PIL import Image, ImageChops, ImageDraw, ImageFilter
    img = img.convert("RGB")
    for op in ops:
        pts = [(x * k, y * k) for x, y in op["pts"]]
        w = max(1.0, op["w"] * k)
        col = _hex_rgb(op["color"])
        kind = op["kind"]
        if kind == "blur":
            (x0, y0), (x1, y1) = pts[0], pts[-1]
            box = (int(max(0, min(x0, x1))), int(max(0, min(y0, y1))),
                   int(min(img.width, max(x0, x1))), int(min(img.height, max(y0, y1))))
            if box[2] - box[0] < 2 or box[3] - box[1] < 2:
                continue
            part = img.crop(box)
            blk = max(5, int(img.width * 0.008))
            small = part.resize((max(1, part.width // blk), max(1, part.height // blk)),
                                Image.BILINEAR)
            part = small.resize(part.size, Image.NEAREST).filter(
                ImageFilter.GaussianBlur(max(2, blk // 2)))
            img.paste(part, box[:2])
            continue
        mask = Image.new("L", img.size, 0)
        d = ImageDraw.Draw(mask)
        if kind == "marker":
            (x0, y0), (x1, y1) = pts[0], pts[-1]
            d.rectangle((min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)),
                        fill=255)
        elif kind == "frame":
            (x0, y0), (x1, y1) = pts[0], pts[-1]
            d.rectangle((min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)),
                        outline=255, width=int(w))
        elif kind == "arrow":
            import math
            (x0, y0), (x1, y1) = pts[0], pts[-1]
            ang = math.atan2(y1 - y0, x1 - x0)
            hl, hw = w * 5, w * 2.6
            bx, by = x1 - hl * math.cos(ang), y1 - hl * math.sin(ang)
            if math.hypot(x1 - x0, y1 - y0) > hl:
                d.line((x0, y0, bx, by), fill=255, width=int(w))
            d.polygon([(x1, y1),
                       (bx + hw * math.sin(ang), by - hw * math.cos(ang)),
                       (bx - hw * math.sin(ang), by + hw * math.cos(ang))], fill=255)
        if kind == "marker":
            # wie ein echter Textmarker: Farbe wird multipliziert, Text bleibt lesbar
            tint = ImageChops.multiply(img, Image.new("RGB", img.size, col))
            img.paste(tint, (0, 0), mask)
            continue
        img.paste(Image.new("RGB", img.size, col), (0, 0), mask)
    return img


def export_all(steps, base_dir: str, title: str, lang: str, fmts: dict,
               restart: bool = True, embed_plain: bool = False):
    """Legt base_dir/<Titel>/ an und erzeugt dort alle gewaehlten Formate
    (HTML-Bilder als PNG in base_dir/<Titel>/img_<Titel>/).
    Rueckgabe: (ordner, [dateien], [(format, fehler)])"""
    name = safe_name(title)
    folder = unique_dir(os.path.join(base_dir, name))
    os.makedirs(folder)
    made, errors = [], []
    # Bilder der HTML-Dateien: Unterordner img_<Titel>/schritt_NN.png
    img_rel = "img_" + name
    img_dir = None if embed_plain else (os.path.join(folder, img_rel), img_rel)
    jobs = [
        ("html", name + ".html",
         lambda p: write_html(steps, p, title, lang, restart, img_dir)),
        ("plain", name + "_einfach.html",
         lambda p: write_plain(steps, p, title, lang, restart, img_dir)),
        ("wiki", name + "_zum_Einfuegen.html",
         lambda p: write_wiki(steps, p, title, lang, restart)),
        ("docx", name + ".docx",
         lambda p: write_docx(steps, p, title, lang, restart)),
        ("pdf", name + ".pdf",
         lambda p: write_pdf(steps, p, title, lang, restart)),
    ]
    for key, fname, fn in jobs:
        if not fmts.get(key):
            continue
        path = os.path.join(folder, fname)
        try:
            fn(path)
            made.append(path)
        except Exception as exc:
            errors.append((key, f"{type(exc).__name__}: {exc}"))
    return folder, made, errors


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
def select_region(root, bounds, hint: str):
    """Vollbild-Overlay ueber alle Monitore; Rechteck aufziehen.
    Gibt (left, top, breite, hoehe) oder None zurueck."""
    import tkinter as tk
    l, t, w, h = bounds
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    top.attributes("-topmost", True)
    top.attributes("-alpha", 0.35)
    top.geometry(f"{w}x{h}{l:+d}{t:+d}")
    cv = tk.Canvas(top, bg="black", highlightthickness=0, cursor="crosshair")
    cv.pack(fill="both", expand=True)
    cv.create_text(500 - l, 40 - t, text=hint,
                   fill="white", font=("Segoe UI", 16))
    st = {"x0": 0, "y0": 0, "rect": None, "res": None}

    def press(e):
        st["x0"], st["y0"] = e.x_root, e.y_root
        st["cx0"], st["cy0"] = e.x, e.y
        st["rect"] = cv.create_rectangle(e.x, e.y, e.x, e.y, outline="red",
                                         width=3)

    def drag(e):
        if st["rect"]:
            cv.coords(st["rect"], st["cx0"], st["cy0"], e.x, e.y)

    def release(e):
        x0, y0, x1, y1 = st["x0"], st["y0"], e.x_root, e.y_root
        rw, rh = abs(x1 - x0), abs(y1 - y0)
        if rw >= 20 and rh >= 20:
            st["res"] = (min(x0, x1), min(y0, y1), rw, rh)
        top.destroy()

    cv.bind("<ButtonPress-1>", press)
    cv.bind("<B1-Motion>", drag)
    cv.bind("<ButtonRelease-1>", release)
    top.bind("<Escape>", lambda e: top.destroy())
    top.focus_force()
    top.grab_set()
    root.wait_window(top)
    return st["res"]


def select_region_preview(root, hint: str):
    """Fuer macOS/Linux: Screenshot des gesamten Desktops in einem normalen
    Fenster anzeigen, dort Rechteck aufziehen (Koordinaten werden umgerechnet,
    funktioniert mit mehreren Monitoren und Retina)."""
    import tkinter as tk
    import mss
    from PIL import Image, ImageTk
    with new_mss(mss) as sct:
        m = sct.monitors[0]
        shot = sct.grab(m)
        img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
    sc = min(root.winfo_screenwidth() * 0.85 / m["width"],
             root.winfo_screenheight() * 0.75 / m["height"], 1.0)
    dw, dh = max(1, int(m["width"] * sc)), max(1, int(m["height"] * sc))
    photo = ImageTk.PhotoImage(img.resize((dw, dh), Image.LANCZOS))
    top = tk.Toplevel(root)
    top.title(hint)
    top.attributes("-topmost", True)
    cv = tk.Canvas(top, width=dw, height=dh, highlightthickness=0,
                   cursor="crosshair")
    cv.pack()
    cv.create_image(0, 0, image=photo, anchor="nw")
    cv.image = photo
    st = {"p0": None, "rect": None, "res": None}

    def press(e):
        st["p0"] = (e.x, e.y)
        st["rect"] = cv.create_rectangle(e.x, e.y, e.x, e.y, outline="red",
                                         width=3)

    def drag(e):
        if st["rect"] and st["p0"]:
            cv.coords(st["rect"], st["p0"][0], st["p0"][1], e.x, e.y)

    def release(e):
        if st["p0"]:
            x0, y0 = st["p0"]
            x1, y1 = max(0, min(e.x, dw)), max(0, min(e.y, dh))
            if abs(x1 - x0) >= 8 and abs(y1 - y0) >= 8:
                st["res"] = (int(m["left"] + min(x0, x1) / sc),
                             int(m["top"] + min(y0, y1) / sc),
                             int(abs(x1 - x0) / sc), int(abs(y1 - y0) / sc))
        top.destroy()

    cv.bind("<ButtonPress-1>", press)
    cv.bind("<B1-Motion>", drag)
    cv.bind("<ButtonRelease-1>", release)
    top.bind("<Escape>", lambda e: top.destroy())
    top.focus_force()
    root.wait_window(top)
    return st["res"]


def open_file(path: str) -> None:
    try:
        if IS_WIN:
            os.startfile(path)   # type: ignore[attr-defined]
        elif IS_MAC:
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


def settings_path() -> str:
    if IS_WIN:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif IS_MAC:
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.path.expanduser("~/.config")
    return os.path.join(base, "KlickDoku", "settings.json")


def load_settings() -> dict:
    import json
    p = settings_path()
    # Einstellungen frueherer Versionen (Ordner "StepsRecorder") weiterverwenden
    for path in (p, os.path.join(os.path.dirname(os.path.dirname(p)),
                                 "StepsRecorder", "settings.json")):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            continue
    return {}


def save_settings(cfg: dict) -> None:
    try:
        import json
        p = settings_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=1)
    except Exception:
        pass


def run_gui(lang: str) -> None:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    T = TEXTS[lang]
    rec = Recorder(lang)
    cfg = load_settings()
    ACC, ACC_H = "#3b82f6", "#2563eb"
    RED, RED_H = "#ef4444", "#dc2626"
    AMBER = "#f59e0b"
    SOFT = ("#e5e7eb", "#2a2f3a")
    SOFT_H = ("#d1d5db", "#363c4a")
    MUT = ("#6b7280", "#9ca3af")
    TXT = ("#111827", "#e5e7eb")

    root = ctk.CTk()
    root.title(T["app"])
    root.attributes("-topmost", True)
    root.resizable(False, False)
    root.update_idletasks()
    exclude_from_capture(root)
    try:   # Fenster-/Taskleisten-Icon
        if IS_WIN:
            ico = resource("assets", "icon.ico")
            root.after(300, lambda: root.iconbitmap(ico))  # CTk setzt sonst eigenes
        else:
            import tkinter as _tk
            root._icon = _tk.PhotoImage(file=resource("assets", "icon_256.png"))
            root.iconphoto(True, root._icon)
    except Exception:
        pass

    import mss
    with new_mss(mss) as _s:
        n_mon = len(_s.monitors) - 1
        bounds = (_s.monitors[0]["left"], _s.monitors[0]["top"],
                  _s.monitors[0]["width"], _s.monitors[0]["height"])
    modes = [("screen", T["m_screen"])]
    if n_mon > 1:
        modes.append(("all", T["m_all"]))
        modes += [(f"monitor:{i}", T["m_mon"].format(n=i))
                  for i in range(1, n_mon + 1)]
    modes += [("window", T["m_window"]), ("target", T["m_target"]),
              ("region", T["m_region"])]
    mode_labels = [m[1] for m in modes]

    fm = cfg.get("formats", {})
    var_fmt = {k: ctk.IntVar(value=int(fm.get(k, True)))
               for k in ("html", "plain", "wiki", "docx", "pdf")}
    var_restart = ctk.IntVar(value=int(cfg.get("restart", True)))
    var_embed = ctk.IntVar(value=int(cfg.get("embed", False)))

    def small(parent, text):
        return ctk.CTkLabel(parent, text=text, anchor="w", text_color=MUT,
                            font=ctk.CTkFont(size=12))

    card = ctk.CTkFrame(root, corner_radius=18)
    card.pack(padx=12, pady=12, fill="both", expand=True)

    # --- Kopf: Statuspunkt, Status, Zaehler, Hell/Dunkel -----------------
    head = ctk.CTkFrame(card, fg_color="transparent")
    head.pack(fill="x", padx=18, pady=(16, 2))
    dot = ctk.CTkLabel(head, text="●", text_color="#9ca3af",
                       font=ctk.CTkFont(size=18))
    dot.pack(side="left")
    status = ctk.CTkLabel(head, text=T["idle"],
                          font=ctk.CTkFont(size=16, weight="bold"))
    status.pack(side="left", padx=(6, 0))
    ctk.CTkButton(head, text="◐", width=30, height=28, fg_color=SOFT,
                  hover_color=SOFT_H, text_color=TXT,
                  command=lambda: ctk.set_appearance_mode(
                      "light" if ctk.get_appearance_mode() == "Dark" else "dark")
                  ).pack(side="right")
    counter = ctk.CTkLabel(head, text=T["count"].format(n=0), text_color=MUT)
    counter.pack(side="right", padx=10)

    # --- Berichtstitel (wird die H1 im Bericht) ---------------------------
    small(card, T["title_lbl"]).pack(fill="x", padx=20, pady=(12, 0))
    title_var = ctk.StringVar()
    ctk.CTkEntry(card, textvariable=title_var, height=38, corner_radius=10,
                 placeholder_text=T["title_ph"], width=330).pack(
        padx=18, pady=(2, 0), fill="x")

    # --- Aufnahmebereich ---------------------------------------------------
    small(card, T["cap"]).pack(fill="x", padx=20, pady=(10, 0))
    mode_menu = ctk.CTkOptionMenu(card, values=mode_labels, height=36,
                                  corner_radius=10, fg_color=SOFT,
                                  button_color=SOFT_H, button_hover_color=SOFT_H,
                                  text_color=TXT,
                                  dropdown_font=ctk.CTkFont(size=13))
    mode_menu.pack(padx=18, pady=(2, 0), fill="x")

    extra = ctk.CTkFrame(card, fg_color="transparent")     # je nach Modus gefuellt
    region_row = ctk.CTkFrame(extra, fg_color="transparent")
    b_region = ctk.CTkButton(region_row, text=T["region_btn"], height=32,
                             corner_radius=10, fg_color="transparent",
                             border_width=1, border_color=SOFT_H, text_color=TXT,
                             hover_color=SOFT)
    b_region.pack(fill="x")
    win_row = ctk.CTkFrame(extra, fg_color="transparent")
    win_menu = ctk.CTkOptionMenu(win_row, values=[T["win_none"]], height=34,
                                 corner_radius=10, fg_color=SOFT,
                                 button_color=SOFT_H, button_hover_color=SOFT_H,
                                 text_color=TXT, dynamic_resizing=False, width=250)
    win_menu.pack(side="left", fill="x", expand=True)
    b_refresh = ctk.CTkButton(win_row, text="↻", width=36, height=34,
                              corner_radius=10, fg_color=SOFT, hover_color=SOFT_H,
                              text_color=TXT)
    b_refresh.pack(side="left", padx=(6, 0))

    # --- Hauptknopf + Zusatzknoepfe ----------------------------------------
    b_start = ctk.CTkButton(card, text="⏺  " + T["start"].replace("&&", "&"),
                            height=48, corner_radius=14, fg_color=ACC,
                            hover_color=ACC_H,
                            font=ctk.CTkFont(size=15, weight="bold"))
    b_start.pack(padx=18, pady=(16, 8), fill="x")

    row = ctk.CTkFrame(card, fg_color="transparent")
    row.pack(padx=18, fill="x")
    sec_kw = dict(height=36, corner_radius=10, fg_color=SOFT, hover_color=SOFT_H,
                  text_color=TXT, state="disabled")
    b_head = ctk.CTkButton(row, text="¶ " + T["heading_btn"], width=10, **sec_kw)
    b_pause = ctk.CTkButton(row, text="⏸ " + T["pause"], width=10, **sec_kw)
    b_comm = ctk.CTkButton(row, text="✎ " + T["comment"], width=10, **sec_kw)
    for c, b in enumerate((b_head, b_pause, b_comm)):
        row.grid_columnconfigure(c, weight=1, uniform="a")
        b.grid(row=0, column=c, padx=3, sticky="ew")

    # --- Ausgabe & Optionen (einklappbar) ----------------------------------
    b_opts = ctk.CTkButton(card, text="⚙  " + T["opts"] + "  ▸", height=28,
                           fg_color="transparent", hover_color=SOFT,
                           text_color=MUT, anchor="w")
    b_opts.pack(padx=14, pady=(10, 0), fill="x")
    opts = ctk.CTkFrame(card, fg_color="transparent")
    cb_kw = dict(font=ctk.CTkFont(size=12), checkbox_width=18, checkbox_height=18)
    for key, label in (("html", T["fmt_html"]), ("plain", T["fmt_plain"]),
                       ("wiki", T["fmt_wiki"]), ("docx", T["fmt_docx"]), ("pdf", T["fmt_pdf"])):
        ctk.CTkCheckBox(opts, text=label, variable=var_fmt[key], **cb_kw).pack(
            anchor="w", pady=2)
    ctk.CTkFrame(opts, height=1, fg_color=SOFT_H).pack(fill="x", pady=6)
    for var, label in ((var_restart, T["opt_restart"]),
                       (var_embed, T["opt_embed"])):
        ctk.CTkCheckBox(opts, text=label, variable=var, **cb_kw).pack(
            anchor="w", pady=2)

    def toggle_opts() -> None:
        if opts.winfo_ismapped():
            opts.pack_forget()
            b_opts.configure(text="⚙  " + T["opts"] + "  ▸")
        else:
            opts.pack(padx=24, pady=(2, 0), fill="x", after=b_opts)
            b_opts.configure(text="⚙  " + T["opts"] + "  ▾")
    b_opts.configure(command=toggle_opts)

    ctk.CTkLabel(card, text=T["hint_mac"] if IS_MAC else T["hint"],
                 text_color=MUT, font=ctk.CTkFont(size=11),
                 justify="center").pack(padx=14, pady=(12, 16))

    # --- Logik -------------------------------------------------------------
    def set_status(text: str, color: str) -> None:
        status.configure(text=text)
        dot.configure(text_color=color)

    def set_state(recording: bool) -> None:
        if recording:
            b_start.configure(text="⏹  " + T["stop"].replace("&&", "&"),
                              fg_color=RED, hover_color=RED_H)
            set_status(T["rec"], RED)
        else:
            b_start.configure(text="⏺  " + T["start"].replace("&&", "&"),
                              fg_color=ACC, hover_color=ACC_H)
            set_status(T["idle"], "#9ca3af")
        st = "normal" if recording else "disabled"
        for b in (b_head, b_pause, b_comm):
            b.configure(state=st)
        b_pause.configure(text="⏸ " + T["pause"])

    def choose_region() -> bool:
        if IS_WIN:
            res = select_region(root, bounds, T["region_hint"])
        else:
            res = select_region_preview(root, T["region_hint"])
        if res:
            rec.region = res
            b_region.configure(text=T["region_btn"] + "  (" + T["region_info"].format(
                w=res[2], h=res[3]) + ")")
            return True
        return False

    win_map: dict = {}

    def select_window(label: str) -> None:
        rec.target = win_map.get(label)

    def refresh_windows() -> None:
        win_map.clear()
        labels = []
        for hid, wtitle, pid in list_windows(os.getpid()):
            lab = wtitle if len(wtitle) <= 46 else wtitle[:45] + "…"
            base, k = lab, 2
            while lab in win_map:
                lab, k = f"{base} ({k})", k + 1
            win_map[lab] = {"id": hid, "pid": pid, "title": wtitle}
            labels.append(lab)
        if not labels:
            labels = [T["win_none"]]
        win_menu.configure(values=labels)
        cur = win_menu.get()
        win_menu.set(cur if cur in win_map else labels[0])
        select_window(win_menu.get())

    win_menu.configure(command=select_window)
    b_refresh.configure(command=refresh_windows)

    def on_mode(choice: str) -> None:
        rec.mode = modes[mode_labels.index(choice)][0]
        extra.pack_forget()
        region_row.pack_forget()
        win_row.pack_forget()
        if rec.mode == "region":
            region_row.pack(fill="x")
            extra.pack(padx=18, pady=(6, 0), fill="x", after=mode_menu)
            if rec.region is None and not choose_region():
                mode_menu.set(mode_labels[0])
                on_mode(mode_labels[0])
        elif rec.mode == "target":
            win_row.pack(fill="x")
            extra.pack(padx=18, pady=(6, 0), fill="x", after=mode_menu)
            refresh_windows()

    mode_menu.configure(command=on_mode)
    mode_menu.set(mode_labels[0])
    b_region.configure(command=choose_region)

    def toggle_pause() -> None:
        if not rec.active:
            return
        rec.paused = not rec.paused
        b_pause.configure(text=("▶ " + T["resume"]) if rec.paused
                          else "⏸ " + T["pause"])
        set_status(T["paused"] if rec.paused else T["rec"],
                   AMBER if rec.paused else RED)

    def ask(title: str, prompt: str) -> str:
        was = rec.paused
        rec.paused = True
        if IS_MAC:
            root.deiconify()
        txt = ctk.CTkInputDialog(text=prompt, title=title).get_input()
        if IS_MAC and rec.active:
            root.iconify()
        rec.paused = was
        return (txt or "").strip()

    def add_comment() -> None:
        if rec.active:
            txt = ask(T["comment_title"], T["comment_prompt"])
            if txt:
                rec.add_comment(txt)

    def add_heading() -> None:
        if rec.active:
            txt = ask(T["heading_title"], T["heading_prompt"])
            if txt:
                rec.add_heading(txt)

    # --- Bildeditor: Textmarker, Rahmen, Pfeil, Verwischen -----------------
    def edit_image(png: bytes):
        """Oeffnet den Editor; Rueckgabe: neue PNG-Bytes oder None."""
        import tkinter as tk
        from tkinter import colorchooser
        from PIL import Image, ImageTk
        base = Image.open(io.BytesIO(png)).convert("RGB")
        W, H = base.size
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        k = min((sw - 120) / W, (sh - 260) / H, 1.0)
        mw, fw = edit_widths(base.size)
        wid = {"marker": mw, "frame": fw, "arrow": fw, "blur": 0}
        cols = {"marker": "#ffe600", "frame": "#e60000", "arrow": "#e60000",
                "blur": "#000000"}
        tools = [("marker", T["ed_marker"]), ("frame", T["ed_frame"]),
                 ("arrow", T["ed_arrow"]), ("blur", T["ed_blur"])]
        ops: list = []
        st = {"tool": "frame", "tmp": None, "start": None, "pts": [], "photo": None,
              "ok": False}
        win = ctk.CTkToplevel(root)
        win.title(T["ed_title"])
        win.attributes("-topmost", True)
        win.geometry(f"{int(W * k) + 40}x{int(H * k) + 130}")
        tb = ctk.CTkFrame(win, fg_color="transparent")
        tb.pack(fill="x", padx=12, pady=(10, 4))
        tool_btns = {}
        sw_frame = ctk.CTkFrame(tb, fg_color="transparent")

        def show_colors():
            for w in sw_frame.winfo_children():
                w.destroy()
            if st["tool"] == "blur":
                return
            ctk.CTkLabel(sw_frame, text=T["ed_color"] + ":",
                         font=ctk.CTkFont(size=12)).pack(side="left", padx=(8, 4))
            for c in ("#e60000", "#ffe600", "#00b050", "#0070f3", "#000000",
                      "#ffffff"):
                b = tk.Button(sw_frame, bg=c, width=2, height=1, relief="solid",
                              bd=2 if cols[st["tool"]] == c else 1,
                              highlightthickness=0,
                              command=lambda c=c: set_color(c))
                b.pack(side="left", padx=2)
            tk.Button(sw_frame, text="…", width=2, height=1, relief="solid", bd=1,
                      command=pick_color).pack(side="left", padx=2)

        def set_color(c):
            cols[st["tool"]] = c
            show_colors()

        def pick_color():
            win.attributes("-topmost", False)
            res_c = colorchooser.askcolor(color=cols[st["tool"]], parent=win)
            win.attributes("-topmost", True)
            if res_c and res_c[1]:
                set_color(res_c[1])

        def set_tool(t):
            st["tool"] = t
            for key, b in tool_btns.items():
                b.configure(fg_color=ACC if key == t else SOFT,
                            text_color="#ffffff" if key == t else TXT)
            show_colors()

        for key, label in tools:
            b = ctk.CTkButton(tb, text=label, width=96, height=34, corner_radius=10,
                              fg_color=SOFT, hover_color=SOFT_H, text_color=TXT,
                              command=lambda key=key: set_tool(key))
            b.pack(side="left", padx=2)
            tool_btns[key] = b
        sw_frame.pack(side="left")
        cv = tk.Canvas(win, width=int(W * k), height=int(H * k), highlightthickness=0,
                       cursor="crosshair", bg="#888888")
        cv.pack(padx=12, pady=6)
        bar2 = ctk.CTkFrame(win, fg_color="transparent")
        bar2.pack(side="bottom", fill="x", padx=12, pady=(0, 10))

        def refresh():
            prev = base.resize((max(1, int(W * k)), max(1, int(H * k))), Image.LANCZOS)
            prev = render_edits(prev, ops, k)
            st["photo"] = ImageTk.PhotoImage(prev)
            cv.delete("all")
            cv.create_image(0, 0, anchor="nw", image=st["photo"])

        def undo():
            if ops:
                ops.pop()
                refresh()

        def apply():
            st["ok"] = True
            win.destroy()

        ctk.CTkButton(bar2, text=T["ed_apply"], height=38, width=140, corner_radius=12,
                      fg_color=ACC, hover_color=ACC_H,
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=apply).pack(side="right")
        ctk.CTkButton(bar2, text=T["review_cancel"], height=38, width=110,
                      corner_radius=12, fg_color=SOFT, hover_color=SOFT_H,
                      text_color=TXT, command=win.destroy).pack(side="right", padx=8)
        ctk.CTkButton(bar2, text="↶ " + T["ed_undo"], height=38, width=130,
                      corner_radius=12, fg_color=SOFT, hover_color=SOFT_H,
                      text_color=TXT, command=undo).pack(side="left")

        def clamp(e):
            return (min(max(e.x, 0), int(W * k)), min(max(e.y, 0), int(H * k)))

        def press(e):
            st["start"] = clamp(e)
            st["pts"] = [st["start"]]

        def drag(e):
            if st["start"] is None:
                return
            x, y = clamp(e)
            t, c = st["tool"], cols[st["tool"]]
            cv.delete("tmp")
            x0, y0 = st["start"]
            if t == "marker":
                cv.create_rectangle(x0, y0, x, y, outline=c, fill=c,
                                    stipple="gray50", tags="tmp")
            elif t == "frame":
                cv.create_rectangle(x0, y0, x, y, outline=c, width=max(2, wid[t] * k),
                                    tags="tmp")
            elif t == "arrow":
                cv.create_line(x0, y0, x, y, fill=c, width=max(2, wid[t] * k),
                               arrow="last", tags="tmp")
            else:
                cv.create_rectangle(x0, y0, x, y, outline="#222222", dash=(4, 3),
                                    tags="tmp")
            st["pts"] = [st["start"], (x, y)]

        def release(e):
            if st["start"] is None:
                return
            drag(e)
            t = st["tool"]
            pts = [(px / k, py / k) for px, py in st["pts"]]
            st["start"] = None
            cv.delete("tmp")
            if len(pts) < 2:
                return
            ops.append(dict(kind=t, color=cols[t], pts=pts, w=wid[t]))
            refresh()

        cv.bind("<ButtonPress-1>", press)
        cv.bind("<B1-Motion>", drag)
        cv.bind("<ButtonRelease-1>", release)
        win.bind("<Control-z>", lambda e: undo())
        set_tool("frame")
        refresh()
        win.after(200, lambda: (win.focus_force(), win.grab_set()))
        root.wait_window(win)
        if not st["ok"] or not ops:
            return None
        out = io.BytesIO()
        render_edits(base, ops, 1.0).save(out, "PNG")
        return out.getvalue()

    # --- Pruefen/Bearbeiten vor dem Speichern -----------------------------
    def review_steps(steps, show_list=True):
        """Dialog vor dem Export: Schritte bearbeiten/einfuegen (show_list) und
        die zu erstellenden Formate waehlen. Rueckgabe: Schritte oder None."""
        from PIL import Image
        work = list(steps)
        res = {"ok": False}
        win = ctk.CTkToplevel(root)
        win.title(T["review_title"] if show_list else T["export_title"])
        win.geometry("880x760" if show_list else "460x330")
        win.attributes("-topmost", True)
        if show_list:
            small(win, T["review_hint"]).pack(pady=(10, 0))
        bar = ctk.CTkFrame(win, fg_color="transparent")
        bar.pack(side="bottom", fill="x", padx=14, pady=12)
        fbox = ctk.CTkFrame(win, corner_radius=10)
        fbox.pack(side="bottom", fill="x", padx=12, pady=(8, 0))
        ctk.CTkLabel(fbox, text=T["fmt_choose"],
                     font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2))
        for key in ("html", "plain", "wiki", "docx", "pdf"):
            ctk.CTkCheckBox(fbox, text=T["fmt_" + key], variable=var_fmt[key],
                            font=ctk.CTkFont(size=12), checkbox_width=18,
                            checkbox_height=18).pack(anchor="w", padx=14, pady=2)
        ctk.CTkFrame(fbox, height=4, fg_color="transparent").pack()
        sc = ctk.CTkScrollableFrame(win)
        if show_list:
            sc.pack(fill="both", expand=True, padx=12, pady=(8, 0))
        thumbs: dict = {}

        def thumb(st):
            if not st.image:
                return None
            if id(st) not in thumbs:
                im = Image.open(io.BytesIO(st.image)).convert("RGB")
                w = 220
                im = im.resize((w, max(1, int(im.height * w / im.width))))
                thumbs[id(st)] = ctk.CTkImage(light_image=im, dark_image=im,
                                              size=im.size)
            return thumbs[id(st)]

        def move(i, d):
            j = i + d
            if 0 <= j < len(work):
                work[i], work[j] = work[j], work[i]
                rebuild()

        def delete(i):
            del work[i]
            rebuild()

        def edit_step_image(st):
            new = edit_image(st.image)
            if new:
                st.image = new
                thumbs.pop(id(st), None)
                rebuild()
            win.after(100, lambda: (win.lift(), win.focus_force()))

        def add_head(i):
            work.insert(i, Step("heading", dt.datetime.now(), text=""))
            rebuild()

        def add_step(i):
            work.insert(i, Step("manual", dt.datetime.now(), text=""))
            rebuild()

        def rebuild():
            try:
                frac = sc._parent_canvas.yview()[0]
            except Exception:
                frac = 0.0
            for w in sc.winfo_children():
                w.destroy()
            n = 0
            for i, st in enumerate(work):
                r = ctk.CTkFrame(sc, corner_radius=10)
                r.pack(fill="x", pady=4, padx=2)
                r.grid_columnconfigure(2, weight=1)
                if st.kind == "heading":
                    tag, col = "¶ " + T["section_lbl"], ACC
                    if var_restart.get():
                        n = 0
                elif st.kind == "comment":
                    tag, col = T["note_lbl"], AMBER
                else:
                    n += 1
                    tag, col = f"{T['step']} {n}", ACC
                ctk.CTkLabel(r, text=tag, width=84, anchor="w", text_color=col,
                             font=ctk.CTkFont(size=13, weight="bold")).grid(
                    row=0, column=0, padx=(10, 4), pady=10, sticky="nw")
                im = thumb(st)
                if im:
                    lbl = ctk.CTkLabel(r, text="", image=im, cursor="hand2")
                    lbl.grid(row=0, column=1, padx=4, pady=8)
                    lbl.bind("<Button-1>", lambda e, st=st: edit_step_image(st))
                v = ctk.StringVar(value=st.text)
                v.trace_add("write", lambda *a, st=st, v=v: setattr(st, "text", v.get()))
                ctk.CTkEntry(r, textvariable=v, height=34,
                             font=ctk.CTkFont(size=13, weight=(
                                 "bold" if st.kind == "heading" else "normal"))
                             ).grid(row=0, column=2, padx=6, sticky="ew")
                bf = ctk.CTkFrame(r, fg_color="transparent")
                bf.grid(row=0, column=3, padx=(2, 8))
                bk = dict(width=30, height=30, fg_color=SOFT, hover_color=SOFT_H,
                          text_color=TXT)
                ctk.CTkButton(bf, text="＋", command=lambda i=i: add_step(i),
                              **bk).pack(side="left", padx=1)
                ctk.CTkButton(bf, text="¶", command=lambda i=i: add_head(i),
                              **bk).pack(side="left", padx=1)
                ctk.CTkButton(bf, text="↑", command=lambda i=i: move(i, -1),
                              **bk).pack(side="left", padx=1)
                ctk.CTkButton(bf, text="↓", command=lambda i=i: move(i, 1),
                              **bk).pack(side="left", padx=1)
                ctk.CTkButton(bf, text="✕", command=lambda i=i: delete(i),
                              **dict(bk, fg_color=("#fee2e2", "#4c1d1d"),
                                     hover_color=("#fecaca", "#6b2222"))
                              ).pack(side="left", padx=1)
            win.update_idletasks()
            try:
                sc._parent_canvas.yview_moveto(frac)
            except Exception:
                pass

        def save():
            res["ok"] = True
            win.destroy()

        ctk.CTkButton(bar, text=T["review_save"], height=40, width=160,
                      corner_radius=12, fg_color=ACC, hover_color=ACC_H,
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=save).pack(side="right")
        ctk.CTkButton(bar, text=T["review_cancel"], height=40, width=120,
                      corner_radius=12, fg_color=SOFT, hover_color=SOFT_H,
                      text_color=TXT, command=win.destroy).pack(side="right", padx=8)
        if show_list:
            ctk.CTkButton(bar, text=T["add_step"], height=40, width=150,
                          corner_radius=12, fg_color=SOFT, hover_color=SOFT_H,
                          text_color=TXT,
                          command=lambda: add_step(len(work))).pack(side="left")
            rebuild()
        win.after(200, lambda: (win.focus_force(), win.grab_set()))
        root.wait_window(win)
        if not res["ok"]:
            return None
        return [s for s in work
                if not (s.kind in ("heading", "manual") and not s.text.strip())]

    def persist() -> None:
        cfg.update(formats={k: bool(v.get()) for k, v in var_fmt.items()},
                   restart=bool(var_restart.get()),
                   embed=bool(var_embed.get()))
        save_settings(cfg)

    def finish() -> None:
        rec.stop()
        set_state(False)
        if IS_MAC:
            root.deiconify()
        steps = list(rec.steps)
        if not steps:
            messagebox.showinfo(T["app"], T["nothing"], parent=root)
            return
        # Dialog immer: Schritte pruefen/ergaenzen (optional) + Formate waehlen
        steps = review_steps(steps, show_list=True)
        if steps is None:
            return
        if not steps:
            messagebox.showinfo(T["app"], T["nothing"], parent=root)
            return
        fmts = {k: bool(v.get()) for k, v in var_fmt.items()}
        if not any(fmts.values()):
            fmts["html"] = True
            var_fmt["html"].set(1)
        title = title_var.get().strip() or (
            f"{T['report_title']} – {dt.date.today():%d.%m.%Y}")
        base = None
        while not base:
            base = filedialog.askdirectory(
                parent=root, title=T["choose_dir"], mustexist=True,
                initialdir=cfg.get("last_dir") or os.path.expanduser("~"))
            if not base and messagebox.askyesno(T["app"], T["discard_q"],
                                                parent=root):
                return
        cfg["last_dir"] = base
        persist()
        try:
            folder, made, errors = export_all(
                steps, base, title, lang, fmts, bool(var_restart.get()),
                bool(var_embed.get()))
        except Exception as exc:
            messagebox.showerror(T["app"], str(exc), parent=root)
            return
        if errors:
            messagebox.showerror(T["app"], T["saved_err"].format(
                e="\n".join(f"{k}: {m}" for k, m in errors)), parent=root)
        open_file(folder)

    def toggle() -> None:
        if rec.active:
            finish()
            return
        if rec.mode == "region" and rec.region is None and not choose_region():
            return
        if rec.mode == "target":
            if rec.target is None:
                set_status(T["pick_window"], AMBER)
                return
            activate_window(rec.target)
            time.sleep(0.3)
        set_status(T["starting"], AMBER)
        root.update_idletasks()
        rec.start()
        set_state(True)
        if IS_MAC:
            root.iconify()   # Fenster soll nicht in den Screenshots sein

    b_start.configure(command=toggle)
    b_pause.configure(command=toggle_pause)
    b_comm.configure(command=add_comment)
    b_head.configure(command=add_heading)

    def poll() -> None:
        while not rec.hotkeys.empty():
            k = rec.hotkeys.get()
            if k == "pause":
                toggle_pause()
            elif k == "comment":
                add_comment()
            elif k == "heading":
                add_heading()
            elif k == "stop" and rec.active:
                finish()
        counter.configure(text=T["count"].format(n=count_steps(rec.steps)))
        root.after(200, poll)

    def on_close() -> None:
        if rec.active:
            rec.stop()
        persist()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.after(200, poll)
    root.after(100, root.focus_force)   # Titelfeld nicht vorab fokussieren
    threading.Thread(target=_warm_imports, daemon=True).start()
    root.mainloop()


# --------------------------------------------------------------------------
def selftest(folder: str, lang: str) -> None:
    """Erzeugt Beispielberichte (alle Formate) ohne GUI/Windows."""
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (1280, 720), (230, 235, 245))
    d = ImageDraw.Draw(img)
    d.rectangle((300, 300, 500, 350), outline=(0, 170, 60), width=3)
    d.ellipse((380, 312, 416, 348), outline=(230, 0, 0), width=4)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    now = dt.datetime.now()
    T = TEXTS[lang]
    pic = buf.getvalue()
    steps = [
        Step("heading", now, text="Datei speichern"),
        Step("click", now, text=T["on_elem"].format(
            b=T["left"], name="Speichern", etype="Button")
            + T["in_win"].format(w="Editor"), image=pic),
        Step("typed", now, text=T["typed"] + T["in_win"].format(w="Editor")),
        Step("comment", now, text="Hier erscheint der Speichern-Dialog."),
        Step("key", now, text=T["keys"].format(k="Strg+S")),
        Step("heading", now, text="Fertig"),
        Step("click", now, text=T["left"] + T["in_win"].format(w="Editor"),
             image=pic),
        Step("key", now, text=T["keys"].format(k="ENTER")),
    ]
    out, made, errs = export_all(
        steps, folder, T["report_title"], lang,
        dict(html=True, plain=True, wiki=True, docx=True, pdf=True))
    print("OK:", out)
    for m in made:
        print("  ", os.path.basename(m))
    for k, m in errs:
        print("FEHLER", k, m)


def main() -> None:
    ap = argparse.ArgumentParser(description="KlickDoku - Schrittaufzeichnung")
    ap.add_argument("--lang", choices=list(TEXTS))
    ap.add_argument("--selftest", metavar="ORDNER",
                    help="Beispielberichte (alle Formate) in ORDNER erzeugen")
    a = ap.parse_args()
    lang = pick_lang(a.lang)
    if a.selftest:
        selftest(a.selftest, lang)
        return
    if not (IS_WIN or IS_MAC):
        print("Hinweis: Fuer Windows und macOS entwickelt; Fensternamen und "
              "UI-Elemente stehen nur dort zur Verfuegung.", file=sys.stderr)
    make_dpi_aware()
    if IS_WIN:   # eigene Taskleistengruppe -> eigenes Icon statt python.exe
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "KlickDoku.App")
        except Exception:
            pass
    run_gui(lang)


if __name__ == "__main__":
    main()
