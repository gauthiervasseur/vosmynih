#!/usr/bin/env python3
"""восьминіг - desktop app: drop a Marp HTML deck, get an editable PowerPoint.

    python app/vosmynih_app.py

The window is an HTML/JS arcade UI (app/ui) driven through pywebview; the conversion itself is
marp_html_to_pptx.convert(), run in a background thread.
"""
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time

import webview

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))          # marp_html_to_pptx.py lives one level up (dev runs)
import marp_html_to_pptx as conv                   # noqa: E402

VERSION = "1.0.0"
APP_NAME = "восьминіг"
RES = getattr(sys, "_MEIPASS", HERE)               # PyInstaller unpacks data files here
UI = os.path.join(RES, "ui")


def play_sound(path: str):
    """Play a WAV without blocking: the octopus greets you at launch."""
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["afplay", path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        elif os.name == "nt":
            import winsound
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            for player in ("paplay", "aplay"):
                if shutil.which(player):
                    subprocess.Popen([player, path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    break
    except Exception:
        pass


def inspect_deck(path: str) -> dict:
    """Cheap look at a deck: is it a Marp export, and how many slides."""
    try:
        with open(path, encoding="utf-8", errors="replace") as fh:
            text = fh.read()
    except OSError as exc:
        return {"error": f"Cannot read the file: {exc.strerror or exc}"}
    slides = len(re.findall(r"<svg[^>]*data-marpit-svg", text))
    if not slides:
        return {"error": "This HTML is not a Marp deck (no Marp slides found)."}
    return {"path": path, "name": os.path.basename(path), "folder": os.path.basename(os.path.dirname(path)) or "/",
            "slides": slides, "size_mb": round(os.path.getsize(path) / 1048576, 1)}


class Api:
    """Methods callable from the UI as window.pywebview.api.<name>()."""

    def __init__(self):
        self.window = None
        self._job = None
        self.selftest = None                            # (deck, output) when run with --selftest

    # ---- events to the UI -------------------------------------------------
    def _push(self, event: dict):
        if self.window:
            self.window.evaluate_js(f"window.octo && window.octo.event({json.dumps(event)})")

    # ---- environment & files ----------------------------------------------
    def env(self):
        e = conv.environment()
        return {"chrome": bool(e["chrome"]), "calibri": e["calibri"], "version": VERSION, "platform": sys.platform}

    def choose_deck(self):
        kind = getattr(webview, "FileDialog", None)
        result = self.window.create_file_dialog(kind.OPEN if kind else webview.OPEN_DIALOG,
                                                file_types=("Marp HTML deck (*.html;*.htm)", "All files (*.*)"))
        if not result:
            return None
        return inspect_deck(result[0] if isinstance(result, (list, tuple)) else result)

    def choose_output(self, deck_path):
        kind = getattr(webview, "FileDialog", None)
        base = os.path.splitext(os.path.basename(deck_path))[0] + ".pptx"
        result = self.window.create_file_dialog(kind.SAVE if kind else webview.SAVE_DIALOG,
                                                directory=os.path.dirname(deck_path), save_filename=base,
                                                file_types=("PowerPoint (*.pptx)",))
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        return path if path.lower().endswith(".pptx") else path + ".pptx"

    def dropped(self, path):
        self._push({"type": "deck", "deck": inspect_deck(path)})

    # ---- conversion ---------------------------------------------------------
    def start(self, opts):
        if self._job and self._job.is_alive():
            return False

        def work():
            t0 = time.time()
            last = [0.0]

            def progress(stage, done, total):
                now = time.time()                       # throttle UI updates during the slide loop
                if stage != "slides" or done == total or now - last[0] > 0.08:
                    last[0] = now
                    self._push({"type": "progress", "stage": stage, "done": done, "total": total})

            try:
                summary = conv.convert(
                    opts["path"], opts.get("output") or None, slides=opts.get("slides") or None,
                    download=opts.get("download", True), video=opts.get("video", True),
                    backgrounds=opts.get("backgrounds", True), natural_wrap=not opts.get("breaks", True),
                    approx_layout=not opts.get("exact", True), progress=progress,
                    log=lambda m: self._push({"type": "log", "text": m.strip()}))
                summary["seconds"] = round(time.time() - t0)
                self._push({"type": "done", "summary": summary})
                self._selftest_end(summary)
            except Exception as exc:                    # shown as GAME OVER
                self._push({"type": "error", "message": str(exc) or exc.__class__.__name__})
                self._selftest_end({"error": str(exc)})

        self._job = threading.Thread(target=work, daemon=True)
        self._job.start()
        return True

    def _selftest_end(self, result):
        if self.selftest:
            time.sleep(1.5)                             # let the STAGE CLEAR screen show
            ui = self.window.evaluate_js("document.getElementById('win').classList.contains('on')")
            print("SELFTEST", json.dumps({**result, "ui_stage_clear": ui}), flush=True)
            self.window.destroy()

    # ---- after the win --------------------------------------------------------
    def open_path(self, path):
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)                          # noqa: S606 - opens the user's own file
        else:
            subprocess.Popen(["xdg-open", path])

    def reveal(self, path):
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        elif os.name == "nt":
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])


def bind(window, api):
    """Runs once the window exists: greet with the octopus, wire drag & drop."""
    play_sound(os.path.join(UI, "octopus.wav"))
    try:
        from webview.dom import DOMEventHandler

        def on_drop(e):
            files = (e.get("dataTransfer") or {}).get("files") or []
            path = files[0].get("pywebviewFullPath") if files else None
            if path:
                api.dropped(path)

        window.dom.document.events.dragover += DOMEventHandler(lambda e: None, True, True)
        window.dom.document.events.drop += DOMEventHandler(on_drop, True, True)
    except Exception:
        pass                                            # older pywebview: click-to-open still works
    if api.selftest:                                    # end-to-end check used by the build scripts
        deck, output = api.selftest
        time.sleep(2)
        api.dropped(deck)
        time.sleep(1)
        api.start({"path": deck, "output": output, "slides": "1-3", "video": False})


def main():
    api = Api()
    if "--selftest" in sys.argv:
        i = sys.argv.index("--selftest")
        api.selftest = (sys.argv[i + 1], sys.argv[i + 2])
    window = webview.create_window(APP_NAME, url=os.path.join(UI, "index.html"), js_api=api,
                                   width=1000, height=780, min_size=(880, 700), background_color="#0b0221")
    api.window = window
    webview.start(bind, (window, api), debug="--debug" in sys.argv)


if __name__ == "__main__":
    main()
