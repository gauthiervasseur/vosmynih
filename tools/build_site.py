#!/usr/bin/env python3
"""Assemble the восьминіг website into site/ (static, ready to upload anywhere).

    python tools/build_site.py

Uses web/index.html as the template, app/ui as the playable demo, and every installer found in
dist/ (DMG, Windows zip / setup.exe) as downloads, plus a source zip built from the repo.
"""
import hashlib
import html
import os
import re
import shutil
import subprocess
import sys
import zipfile

from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE = os.path.join(ROOT, "site")
sys.path.insert(0, ROOT)
from marp_html_to_pptx import find_chrome   # noqa: E402

VERSION = re.search(r'^VERSION = "(.*)"', open(os.path.join(ROOT, "app", "vosmynih_app.py"), encoding="utf-8").read(), re.M).group(1)
SOURCE_FILES = ["marp_html_to_pptx.py", "requirements.txt", "requirements-app.txt", "readme.txt", ".gitignore",
                "app", "tools", "build/build_macos.sh", "build/build_windows.ps1", "build/installer.iss",
                "build/github-actions-build.yml", "web"]


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def mb(path):
    return f"{os.path.getsize(path) / 1048576:.0f} MB" if os.path.getsize(path) > 1048576 else f"{os.path.getsize(path) / 1024:.0f} KB"


def source_zip(dest):
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for item in SOURCE_FILES:
            full = os.path.join(ROOT, item)
            paths = [full] if os.path.isfile(full) else [os.path.join(d, f) for d, _, fs in os.walk(full) for f in fs]
            for p in paths:
                if "__pycache__" in p or p.endswith(".DS_Store"):
                    continue
                z.write(p, os.path.join(f"vosmynih-{VERSION}", os.path.relpath(p, ROOT)))


def screenshots(dest_dir):
    chrome = find_chrome(None)
    if not chrome:
        print("  ! Chrome not found: no screenshots")
        return
    ui = os.path.join(ROOT, "app", "ui", "index.html")
    for state in ("loaded", "win"):
        png = os.path.join(dest_dir, f"shot-{state}.png")
        prof = os.path.join(dest_dir, f".prof-{state}")
        subprocess.run([chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", f"--user-data-dir={prof}",
                        "--window-size=1000,780", "--virtual-time-budget=4500", f"--screenshot={png}",
                        f"file://{ui}?state={state}"], capture_output=True)
        shutil.rmtree(prof, ignore_errors=True)
        if os.path.exists(png):
            Image.open(png).convert("RGB").save(png, optimize=True)


def card(title, meta, href, label, size, digest):
    return (f'      <div class="card"><h3>{html.escape(title)}</h3><p class="meta">{html.escape(meta)}<br>SHA-256: {digest[:16]}…</p>\n'
            f'        <a class="btn main" href="{href}" download>{html.escape(label)}<small>{size}</small></a></div>')


def main():
    shutil.rmtree(SITE, ignore_errors=True)
    assets, downloads = os.path.join(SITE, "assets"), os.path.join(SITE, "downloads")
    os.makedirs(assets)
    os.makedirs(downloads)
    ui = os.path.join(ROOT, "app", "ui")
    shutil.copytree(ui, os.path.join(SITE, "play"))                 # the real UI, in demo mode
    for f in ("sprite.js", "octopus.wav", "PressStart2P-Regular.ttf", "PressStart2P-OFL.txt"):
        shutil.copy(os.path.join(ui, f), assets)
    icon = Image.open(os.path.join(ROOT, "app", "assets", "icon.png"))
    icon.resize((64, 64), Image.LANCZOS).save(os.path.join(assets, "icon-64.png"))
    icon.resize((32, 32), Image.LANCZOS).save(os.path.join(assets, "favicon.png"))
    icon.resize((256, 256), Image.LANCZOS).save(os.path.join(assets, "icon-256.png"))
    screenshots(assets)

    cards, platforms = [], []
    dist = os.path.join(ROOT, "dist")
    found = sorted(os.listdir(dist)) if os.path.isdir(dist) else []
    for name in found:
        path = os.path.join(dist, name)
        m = re.match(r"vosmynih-[\d.]+-macos-(arm64|x86_64)\.dmg$", name)
        if m:
            shutil.copy(path, downloads)
            arch = "Apple Silicon (M1-M4)" if m.group(1) == "arm64" else "Intel"
            platforms.append("MAC")
            cards.append(card(f"MAC · {arch.upper()}", f"{name} · macOS 11 or later", f"downloads/{name}", "▶ DOWNLOAD .DMG", mb(path), sha256(path)))
        elif re.match(r"vosmynih-[\d.]+-windows-x64(-setup\.exe|\.zip)$", name):
            shutil.copy(path, downloads)
            kind = "SETUP.EXE" if name.endswith(".exe") else "PORTABLE .ZIP"
            platforms.append("WINDOWS")
            cards.append(card(f"WINDOWS · {kind}", f"{name} · Windows 10 or 11", f"downloads/{name}", f"▶ DOWNLOAD {kind}", mb(path), sha256(path)))
    if not any("WINDOWS" in c for c in cards):
        cards.append('      <div class="card"><h3>WINDOWS</h3><p class="meta">Build the installer on a Windows PC with '
                     '<code>build\\build_windows.ps1</code> (in the source download). It makes a setup.exe and a portable zip.</p>'
                     '<div class="soon">INSTALLER<br>BUILT FROM SOURCE</div></div>')
    src = os.path.join(downloads, f"vosmynih-{VERSION}-source.zip")
    source_zip(src)
    cards.append(card("SOURCE CODE", f"{os.path.basename(src)} · Python app, converter, build scripts",
                      f"downloads/{os.path.basename(src)}", "▶ DOWNLOAD SOURCE", mb(src), sha256(src)))

    page = open(os.path.join(ROOT, "web", "index.html"), encoding="utf-8").read()
    page = (page.replace("{{VERSION}}", VERSION).replace("{{DOWNLOADS}}", "\n".join(cards))
                .replace("{{PLATFORMS}}", " · ".join(dict.fromkeys(platforms)) or "SOURCE CODE"))
    with open(os.path.join(SITE, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(page)
    total = sum(os.path.getsize(os.path.join(d, f)) for d, _, fs in os.walk(SITE) for f in fs)
    print(f"site/ ready ({total / 1048576:.0f} MB): open site/index.html, or upload the folder to any static host")
    for name in sorted(os.listdir(downloads)):
        print("  downloads/" + name, mb(os.path.join(downloads, name)))


if __name__ == "__main__":
    main()
