#!/usr/bin/env python3
"""
marp_html_to_pptx.py - convert a Marp HTML export into an EDITABLE PowerPoint deck that keeps
the HTML's alignment.

Headless Chrome lays out every slide with the deck's own CSS and reports, for each element,
its box, font size, line height, colours, alignment, first-line position and line breaks.
Each element is then rebuilt from native PowerPoint objects at those coordinates:
  - headings, paragraphs, bullet / numbered lists -> text boxes (same line breaks as the HTML)
  - tables -> native tables (cell sizes, fills, borders from the CSS)
  - code blocks -> monospace text boxes, syntax colours kept
  - images -> pictures with object-fit (contain / cover), clipped to the slide like the HTML
  - MathJax equations -> PNG pictures (needs cairosvg)
  - videos -> embedded movies (or a placeholder with --no-video)
  - CSS backgrounds/borders of blocks (cards, quotes) -> shapes
  - theme backgrounds (colour, mosaic, wordmark, rule) -> real slide backgrounds
  - speaker notes -> PowerPoint notes
Chrome measures text with Calibri taken from the local Office install, so PowerPoint draws the
lines where the browser did. Without Chrome, a built-in approximate layout is used instead.

Usage:
    pip install python-pptx beautifulsoup4 lxml pillow cairosvg
    python marp_html_to_pptx.py deck.html                 # -> deck.pptx next to the HTML
    python marp_html_to_pptx.py deck.html -o out.pptx --slides 1-40,120
    python marp_html_to_pptx.py deck.html --no-download --no-video
    python marp_html_to_pptx.py deck.html --natural-wrap          # PowerPoint wraps lines itself
    python marp_html_to_pptx.py deck.html --approx-layout --no-backgrounds   # no Chrome at all
"""
from __future__ import annotations

import argparse
import base64
import html as htmllib
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
import urllib.parse
import urllib.request
from dataclasses import dataclass, replace

from bs4 import BeautifulSoup, Comment, NavigableString, Tag
from lxml import etree
from PIL import Image
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Pt

try:
    import cairosvg  # optional: renders MathJax SVG equations
except Exception:  # cairosvg needs the cairo C library
    cairosvg = None

_LOG = None   # optional callable(str) set by convert(); warnings go to stderr otherwise


def warn(msg: str):
    if _LOG:
        _LOG(msg)
    else:
        print(msg, file=sys.stderr)


NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # no console flashes on Windows


# --------------------------------------------------------------------------
# Geometry: Marp renders on a 1280x720 CSS-px canvas = 13.333" x 7.5"
# --------------------------------------------------------------------------
CANVAS_W, CANVAS_H = 1280, 720
PAD_T, PAD_X, PAD_B = 44, 56, 84
CONTENT_W = CANVAS_W - 2 * PAD_X
BODY_BOTTOM = CANVAS_H - PAD_B
GAP = 28
SCALES = [1.0, 0.93, 0.86, 0.8, 0.74, 0.68, 0.62, 0.56, 0.5]


def E(px: float) -> Emu:
    """CSS px -> EMU (96 dpi)."""
    return Emu(int(round(px * 9525)))


def PT(px: float) -> Pt:
    """CSS px -> points."""
    return Pt(round(px * 0.75, 1))


# --------------------------------------------------------------------------
# Theme (values taken from the deck's "stanford" Marp theme CSS)
# --------------------------------------------------------------------------
SU = {
    "cardinal": "8C1515", "black": "2E2D29", "fog": "E6E6E6", "fog_light": "F4F4F4",
    "cool_grey": "53565A", "poppy": "E98300", "palo_alto": "175E54",
    "illuminating": "FEC51D", "white": "FFFFFF",
}
PALETTES = {
    1: dict(bg="FFFFFF", fg=SU["black"], heading=SU["cardinal"], accent=SU["cardinal"], card_bg=SU["fog_light"], border="B6B1A9"),
    2: dict(bg=SU["cardinal"], fg="FFFFFF", heading="FFFFFF", accent=SU["illuminating"], card_bg="741111", border="B35B5B"),
    3: dict(bg=SU["poppy"], fg=SU["black"], heading=SU["black"], accent=SU["black"], card_bg="FDF0E0", border="B86700"),
    4: dict(bg="FFFFFF", fg=SU["black"], heading=SU["cardinal"], accent=SU["cardinal"], card_bg="FFFFFF", border="B6B1A9"),
    5: dict(bg=SU["black"], fg="FFFFFF", heading="FFFFFF", accent="FFFFFF", card_bg="45433E", border="7B7973"),
    6: dict(bg="E2F0EE", fg=SU["black"], heading=SU["palo_alto"], accent=SU["palo_alto"], card_bg="FFFFFF", border="8CC2BA"),
    7: dict(bg=SU["illuminating"], fg=SU["black"], heading=SU["cardinal"], accent=SU["cardinal"], card_bg="FFF6D6", border="C99A0E"),
    8: dict(bg=SU["palo_alto"], fg="FFFFFF", heading="FFFFFF", accent=SU["illuminating"], card_bg="0F4A42", border="4F8A81"),
}
RULE_COLOR = "8D8B87"
HEAD_FONT, BODY_FONT, CODE_FONT = "Georgia", "Calibri", "Courier New"
CHAR_W = {HEAD_FONT: 0.55, BODY_FONT: 0.49, CODE_FONT: 0.60}  # avg glyph width / font size


def luminance(hexc: str) -> float:
    r, g, b = (int(hexc[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def text_on(hexc: str) -> str:
    return "FFFFFF" if luminance(hexc) < 0.55 else SU["black"]


def rgb(hexc: str) -> RGBColor:
    return RGBColor.from_string(hexc)


# --------------------------------------------------------------------------
# Content model
# --------------------------------------------------------------------------
@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    code: bool = False
    href: str | None = None
    br: bool = False
    color: str | None = None      # exact-layout runs carry their own computed style
    size: float | None = None     # CSS px
    font: str | None = None
    underline: bool = False
    spacing: float | None = None  # letter-spacing, CSS px


@dataclass
class Para:
    runs: list
    size: float                      # CSS px at scale 1
    font: str = BODY_FONT
    color: str | None = None         # None -> palette fg
    bold: bool = False
    bullet: tuple | None = None      # ("ul", level) | ("ol", level, start)
    space_after: float = 0.45        # in em
    line: float = 1.12
    align: str = "left"

    def segments(self):
        """Text split on explicit line breaks (for height estimation)."""
        segs, cur = [], ""
        for r in self.runs:
            if r.br:
                segs.append(cur)
                cur = ""
            else:
                cur += r.text
        segs.append(cur)
        return segs


def clean_runs(runs: list) -> list:
    """Collapse HTML whitespace across runs and trim around line breaks."""
    out = []
    for r in runs:
        if r.br:
            if out and not out[-1].br:
                out[-1].text = out[-1].text.rstrip()
            if out:
                out.append(r)
            continue
        t = r.text
        if not out or out[-1].br or out[-1].text.endswith(" "):
            t = t.lstrip()
        if t:
            r.text = t
            out.append(r)
    while out:
        if out[-1].br:
            out.pop()
            continue
        out[-1].text = out[-1].text.rstrip()
        if out[-1].text:
            break
        out.pop()
    return [r for r in out if r.br or r.text]


def apply_breaks(runs: list, breaks: list | None) -> list:
    """Insert a line break before the words where the browser wrapped.

    `breaks` counts non-whitespace characters, which makes the match independent of how
    each HTML parser split whitespace into text nodes."""
    if not breaks:
        return runs
    todo = sorted(set(breaks))
    out, count, bi = [], 0, 0
    for r in runs:
        if r.br:
            out.append(r)
            continue
        buf = ""
        for ch in r.text:
            if not ch.isspace():
                while bi < len(todo) and todo[bi] < count:
                    bi += 1
                if bi < len(todo) and todo[bi] == count:
                    bi += 1
                    seg = buf.rstrip()
                    if seg:
                        out.append(replace(r, text=seg))
                    elif out and not out[-1].br:
                        out[-1].text = out[-1].text.rstrip()
                    if not (out and out[-1].br):      # the source already breaks here (<br>, new <p>)
                        out.append(Run(br=True))
                    buf = ""
                count += 1
            elif not buf and out and out[-1].br:
                continue                             # no leading space on the new line
            buf += ch
        if buf:
            out.append(replace(r, text=buf))
    return [r for r in out if r.br or r.text]


def runs_text(runs: list) -> str:
    return "".join("\n" if r.br else r.text for r in runs)


# --------------------------------------------------------------------------
# Images
# --------------------------------------------------------------------------
class ImageStore:
    NATIVE = {"PNG", "JPEG", "GIF", "BMP", "TIFF"}

    def __init__(self, html_dir: str, download: bool, max_px: int):
        self.html_dir = html_dir
        self.download = download
        self.max_px = max_px
        self.cache: dict = {}
        self.missing: list = []

    def _read(self, src: str) -> bytes | None:
        if src.startswith("data:"):
            header, _, payload = src.partition(",")
            return base64.b64decode(payload) if ";base64" in header else urllib.parse.unquote(payload).encode()
        if src.startswith(("http://", "https://")):
            if not self.download:
                return None
            try:
                req = urllib.request.Request(src, headers={"User-Agent": "Mozilla/5.0 marp2pptx"})
                with urllib.request.urlopen(req, timeout=20) as resp:
                    return resp.read()
            except Exception as exc:
                warn(f"  ! download failed: {src[:80]} ({exc})")
                return None
        path = urllib.parse.unquote(src.split("#")[0].split("?")[0])
        if path.startswith("file://"):
            path = path[7:]
        candidates = [path] if os.path.isabs(path) else []
        candidates += [os.path.join(self.html_dir, path), os.path.join(self.html_dir, path.lstrip("/"))]
        for c in candidates:
            c = os.path.normpath(c)
            if os.path.isfile(c):
                with open(c, "rb") as fh:
                    return fh.read()
        return None

    def _normalise(self, data: bytes):
        """Return (bytes, width, height) in a format PowerPoint embeds natively."""
        head = data[:512].lstrip()
        if head.startswith(b"<") and b"<svg" in data[:4096]:
            if cairosvg is None:
                return None
            data = cairosvg.svg2png(bytestring=data, scale=2)
        img = Image.open(io.BytesIO(data))
        fmt, (w, h) = img.format, img.size
        if fmt == "GIF":                      # keep animation untouched
            return data, w, h
        too_big = self.max_px and max(w, h) > self.max_px
        if fmt in self.NATIVE and not too_big:
            return data, w, h
        if too_big:
            img.thumbnail((self.max_px, self.max_px), Image.LANCZOS)
        buf = io.BytesIO()
        if fmt == "JPEG" and img.mode in ("RGB", "L"):
            img.save(buf, "JPEG", quality=88)
        else:
            img.save(buf, "PNG", optimize=True)
        return buf.getvalue(), img.size[0], img.size[1]

    def get(self, src: str):
        if src in self.cache:
            return self.cache[src]
        result = None
        data = self._read(src)
        if data:
            try:
                result = self._normalise(data)
            except Exception as exc:
                warn(f"  ! unreadable image {src[:80]} ({exc})")
        if result is None:
            self.missing.append(src)
        self.cache[src] = result
        return result


# --------------------------------------------------------------------------
# Theme backgrounds: render each distinct slide style with headless Chrome
# --------------------------------------------------------------------------
CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
    "google-chrome", "chromium", "chromium-browser", "microsoft-edge",
] + [os.path.join(os.environ.get(v, ""), sub) for v in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")
     for sub in ("Google/Chrome/Application/chrome.exe", "Microsoft/Edge/Application/msedge.exe") if os.environ.get(v)]
SKIP_ATTRS = {"id", "lang", "data-footer", "data-m2p-slide", "data-marpit-fragments", "data-marpit-pagination",
              "data-marpit-pagination-total"}


def find_chrome(explicit: str | None) -> str | None:
    for c in ([explicit] if explicit else []) + [os.environ.get("CHROME", "")] + CHROME_CANDIDATES:
        if c and (os.path.isfile(c) or shutil.which(c)):
            return c if os.path.isfile(c) else shutil.which(c)
    return None


def section_attrs(sec: Tag) -> dict:
    """Attributes that can influence the section's look (footer text etc. removed)."""
    out = {}
    for k, v in sec.attrs.items():
        if k in SKIP_ATTRS:
            continue
        v = " ".join(v) if isinstance(v, list) else v
        if k == "style":
            v = re.sub(r"--footer:[^;]*;?", "", v)
        out[k] = v
    return out


def section_key(sec: Tag) -> str:
    return json.dumps(section_attrs(sec), sort_keys=True)


def css_hex(value: str | None) -> str | None:
    m = re.match(r"rgba?\(([^)]+)\)", value or "")
    if not m:
        return None
    parts = [float(v) for v in m.group(1).replace("/", ",").split(",")]
    if len(parts) > 3 and parts[3] < 0.5:
        return None
    return "".join(f"{int(round(v)):02X}" for v in parts[:3])


PROBE_JS = """
const out = {};
document.querySelectorAll('section[data-probe]').forEach(sec => {
  const c = sel => { const el = sel ? sec.querySelector(sel) : sec;
                     return el ? getComputedStyle(el)[sel ? 'color' : 'backgroundColor'] : null; };
  out[sec.dataset.probe] = {bg: c(null), fg: c('p'), heading: c('h1'), h2: c('h2'), h3: c('h3'),
    accent: c('.pa'), card_bg: c('.pc'), border: c('.pb'), footer: c('footer')};
});
document.getElementById('probe-out').textContent = JSON.stringify(out);
"""


class ThemeRenderer:
    """Screenshots an empty copy of each distinct section style, and reads its computed colours."""

    def __init__(self, chrome: str, css: str, workers: int = 6):
        self.chrome = chrome
        self.css = css
        self.workers = workers
        self.tmp = tempfile.mkdtemp(prefix="marp2pptx_")

    def _page(self, sections_html: str, extra: str = "") -> str:
        return ("<!DOCTYPE html><html><head><meta charset='utf-8'><style>" + self.css + "</style>"
                "<style>html,body{margin:0;padding:0;overflow:hidden;background:#fff}</style></head><body>"
                + sections_html + extra + "</body></html>")

    @staticmethod
    def _wrap(attrs: dict, inner: str = "", probe: str | None = None) -> str:
        a = " ".join(f'{k}="{htmllib.escape(v, quote=True)}"' for k, v in attrs.items())
        if probe is not None:
            a += f' data-probe="{probe}"'
        return ('<div id=":$p"><svg data-marpit-svg="" viewBox="0 0 1280 720" width="1280" height="720">'
                f'<foreignObject width="1280" height="720"><section {a}>{inner}</section>'
                '</foreignObject></svg></div>')

    def _chrome(self, args: list, tag: str) -> subprocess.CompletedProcess:
        profile = os.path.join(self.tmp, "profile_" + tag)
        cmd = [self.chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
               "--no-default-browser-check", f"--user-data-dir={profile}"] + args
        return subprocess.run(cmd, capture_output=True, text=True, timeout=90, creationflags=NO_WINDOW)

    def _shot(self, i: int, attrs: dict) -> bytes | None:
        page = os.path.join(self.tmp, f"bg_{i}.html")
        png = os.path.join(self.tmp, f"bg_{i}.png")
        with open(page, "w", encoding="utf-8") as fh:
            fh.write(self._page(self._wrap(attrs)))
        self._chrome(["--force-device-scale-factor=2", "--window-size=1280,720", "--virtual-time-budget=2000",
                      f"--screenshot={png}", Path(page).as_uri()], str(i))
        if not os.path.isfile(png):
            return None
        img = Image.open(png).convert("RGB")
        if img.size != (2560, 1440):
            img = img.crop((0, 0, min(img.width, 2560), min(img.height, 1440))).resize((2560, 1440))
        buf = io.BytesIO()
        img.save(buf, "PNG", optimize=True)
        return buf.getvalue()

    def render(self, sections: list) -> dict:
        styles = {}
        for sec in sections:
            styles.setdefault(section_key(sec), section_attrs(sec))
        keys = list(styles)
        warn(f"Rendering {len(keys)} theme backgrounds with headless Chrome ...")

        # 1. computed colours for every style, in a single page
        probe_inner = ("<h1>x</h1><h2>x</h2><h3>x</h3><p>x</p><span class='pa' style='color:var(--accent)'>x</span>"
                       "<span class='pc' style='color:var(--card-bg)'>x</span>"
                       "<span class='pb' style='color:var(--card-border)'>x</span><footer>x</footer>")
        probe_page = os.path.join(self.tmp, "probe.html")
        with open(probe_page, "w", encoding="utf-8") as fh:
            fh.write(self._page("".join(self._wrap(styles[k], probe_inner, str(i)) for i, k in enumerate(keys)),
                                "<pre id='probe-out'></pre><script>" + PROBE_JS + "</script>"))
        colours = {}
        try:
            dom = self._chrome(["--virtual-time-budget=2000", "--dump-dom", Path(probe_page).as_uri()], "probe").stdout
            m = re.search(r"<pre id=\"probe-out\">(.*?)</pre>", dom, re.S)
            colours = json.loads(htmllib.unescape(m.group(1))) if m else {}
        except Exception as exc:
            warn(f"  ! colour probe failed ({exc})")

        # 2. one screenshot per style, in parallel
        with ThreadPoolExecutor(self.workers) as pool:
            shots = list(pool.map(lambda ik: self._shot(ik[0], styles[ik[1]]), enumerate(keys)))

        result = {}
        for i, k in enumerate(keys):
            pal = {name: css_hex(v) for name, v in colours.get(str(i), {}).items()}
            result[k] = {"png": shots[i], "pal": {n: v for n, v in pal.items() if v}}
        shutil.rmtree(self.tmp, ignore_errors=True)
        failed = sum(1 for r in result.values() if r["png"] is None)
        if failed:
            warn(f"  ! {failed} background(s) could not be rendered; flat colour used instead")
        return result


# --------------------------------------------------------------------------
# Exact layout: let Chrome lay out every slide and report each element's box
# --------------------------------------------------------------------------
FONT_DIRS = [
    "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
    "/Applications/Microsoft Excel.app/Contents/Resources/DFonts",
    os.path.expanduser("~/Library/Fonts"), "/Library/Fonts", "C:/Windows/Fonts",
    os.path.expanduser("~/.fonts"), "/usr/share/fonts/truetype/msttcorefonts",
]
# CSS family -> {(weight, style): file}. Chrome measures with the exact fonts PowerPoint will use.
# (Office's Consolas is rejected by Chrome's font sanitizer, so code is measured with the browser's
# monospace fallback; line breaks and vertical positions still come from Chrome.)
OFFICE_FONTS = {
    "Calibri": {(400, "normal"): "calibri.ttf", (700, "normal"): "calibrib.ttf",
                (400, "italic"): "calibrii.ttf", (700, "italic"): "calibriz.ttf"},
}
# CSS family -> font name written into the .pptx
FONT_ALIASES = {"cascadia code": "Consolas", "source sans 3": "Calibri", "source serif 4": "Georgia",
                "serif": "Georgia", "sans-serif": "Calibri", "monospace": "Consolas", "system-ui": "Calibri"}
INLINE_DISPLAYS = {"inline", "inline-block", "inline-flex", "inline-grid", "contents", "ruby"}

# Chrome: ascent / content-area height of the fonts it measures with (win metrics)
CHROME_ASCENT = {"Calibri": 0.780, "Georgia": 0.807, "Consolas": 0.786}
# PowerPoint, measured on its PDF export: with percentage line spacing a line is 1.2 x size x pct
# high and, for any pct other than exactly 100% (which uses per-font placement and is avoided),
# the first baseline sits at a fixed fraction of the line height, whatever the font.
PPT_SINGLE = 1.2
PPT_BASELINE_TIGHT, PPT_BASELINE_LOOSE = 0.758, 0.747     # pct < 1, pct > 1
PPT_FONT_DROP = {"Georgia": 0.038}   # em: PowerPoint sets Georgia's baseline this much deeper (fitted on slides)


def para_font(p: dict):
    """(font, size px, line height px) that drive a paragraph's first line in PowerPoint."""
    sized = [r for r in p["runs"] if not r.br and r.size]
    big = max(sized, key=lambda r: r.size) if sized else None
    return (big.font if big else BODY_FONT), (big.size if big else p["size"]), p["lh"]


def ppt_pct(size: float, lh: float) -> float:
    pct = lh / (PPT_SINGLE * size)
    return 1.004 if abs(pct - 1) < 0.004 else pct      # exactly 100% follows other rules


def ppt_baseline(font: str, size: float, lh: float) -> float:
    """Distance from a PowerPoint paragraph's top to its first baseline (px)."""
    pct = ppt_pct(size, lh)
    return ((PPT_BASELINE_TIGHT if pct < 1 else PPT_BASELINE_LOOSE) * pct * PPT_SINGLE * size
            + PPT_FONT_DROP.get(font, 0.0) * size)


def css_baseline(rec: dict, font: str) -> float | None:
    """Absolute y of the first baseline Chrome drew in this element (px)."""
    ft = rec.get("ft")
    return ft[0] + ft[1] * CHROME_ASCENT.get(font, 0.78) if ft else None

MEASURE_JS = r"""
async function m2pRaster(svg, w, h) {            // MathJax SVG -> PNG data URL, no cairo needed
  const clone = svg.cloneNode(true);
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('width', w); clone.setAttribute('height', h); clone.removeAttribute('style');
  const xml = new XMLSerializer().serializeToString(clone).replace(/currentColor/g, getComputedStyle(svg).color);
  const img = new Image();
  img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(xml);
  await img.decode();
  const c = document.createElement('canvas'); c.width = Math.ceil(w * 3); c.height = Math.ceil(h * 3);
  const ctx = c.getContext('2d'); ctx.scale(3, 3); ctx.drawImage(img, 0, 0, w, h);
  return c.toDataURL('image/png');
}
async function m2pMeasure() {
  const out = {};
  const pending = [];
  const num = v => Math.round((parseFloat(v) || 0) * 100) / 100;
  const INLINE = new Set(['inline', 'inline-block', 'inline-flex', 'inline-grid', 'contents', 'ruby']);
  const firstText = el => document.createTreeWalker(el, NodeFilter.SHOW_TEXT,
      {acceptNode: n => n.nodeValue.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT}).nextNode();
  // Where the browser wraps lines: positions counted in non-whitespace characters of the element's own text
  const SKIP = new Set(['UL', 'OL', 'TABLE', 'PRE', 'MJX-CONTAINER', 'STYLE', 'SCRIPT']);
  const rg = document.createRange();
  const lineBreaks = el => {
    const tw = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {acceptNode: n => {
      for (let p = n.parentElement; p && p !== el; p = p.parentElement) if (SKIP.has(p.tagName)) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT; }});
    const out = []; let count = 0, prev = null, n;
    while ((n = tw.nextNode())) {
      const re = /\S+/g; let m;
      while ((m = re.exec(n.nodeValue))) {
        rg.setStart(n, m.index); rg.setEnd(n, m.index + 1);
        const q = rg.getClientRects()[0];
        if (!q) continue;
        if (prev && q.top > prev.top + prev.height * 0.6) out.push(count);
        prev = q; count += [...m[0]].length;   // code points, like Python (emoji = 1)
      }
    }
    return out;
  };
  document.querySelectorAll('section[data-m2p-slide]').forEach(sec => {
    const sr = sec.getBoundingClientRect();
    sec.querySelectorAll('[data-m2p]').forEach(el => {
      const cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden') return;
      const r = el.getBoundingClientRect();
      if (r.width < 0.5 && r.height < 0.5) return;
      const ow = el.offsetWidth;
      const rec = {x: num(r.left - sr.left), y: num(r.top - sr.top), w: num(r.width), h: num(r.height),
        s: ow ? num(r.width / ow) : 1, d: cs.display, fs: num(cs.fontSize), lh: cs.lineHeight,
        c: cs.color, ta: cs.textAlign, fw: cs.fontWeight, fi: cs.fontStyle, ff: cs.fontFamily,
        bg: cs.backgroundColor, tt: cs.textTransform, td: cs.textDecorationLine, va: cs.verticalAlign,
        lsp: cs.letterSpacing,
        p: [cs.paddingTop, cs.paddingRight, cs.paddingBottom, cs.paddingLeft].map(num),
        b: [cs.borderTopWidth, cs.borderRightWidth, cs.borderBottomWidth, cs.borderLeftWidth].map(num)};
      if (rec.b.some(v => v > 0)) rec.bc = [cs.borderTopColor, cs.borderRightColor, cs.borderBottomColor, cs.borderLeftColor];
      const rad = num(cs.borderTopLeftRadius); if (rad) rec.br = rad;
      if (!INLINE.has(cs.display)) {          // where the first line of text really sits
        const t = firstText(el);
        if (t) {
          const rg = document.createRange(); rg.selectNodeContents(t);
          const q = [...rg.getClientRects()].find(q => q.width > 0);
          if (q) rec.ft = [num(q.top - sr.top), num(q.height)];
          const lb = lineBreaks(el); if (lb.length) rec.lb = lb;
        }
      }
      if (el.tagName === 'LI') { rec.mk = getComputedStyle(el, '::marker').color; rec.ls = cs.listStyleType; }
      if (el.tagName === 'IMG') { rec.nw = el.naturalWidth; rec.nh = el.naturalHeight; rec.of = cs.objectFit; }
      if (el.tagName.toLowerCase() === 'svg' && el.closest('mjx-container'))
        pending.push(m2pRaster(el, r.width, r.height).then(u => { rec.png = u; }).catch(() => {}));
      out[el.dataset.m2p] = rec;
    });
    out['slide' + sec.dataset.m2pSlide] = {bg: getComputedStyle(sec).backgroundColor};
  });
  await Promise.all(pending);
  document.getElementById('m2p-out').textContent = JSON.stringify(out);
}
window.addEventListener('load', () => document.fonts.ready.then(m2pMeasure));
"""


def office_font_css() -> str:
    """@font-face rules (data URIs) so Chrome lays text out with Calibri/Consolas like PowerPoint does."""
    index = {}
    for d in FONT_DIRS:
        if os.path.isdir(d):
            for f in os.listdir(d):
                index.setdefault(f.lower(), os.path.join(d, f))
    rules = []
    for family, variants in OFFICE_FONTS.items():
        for (weight, style), fname in variants.items():
            path = index.get(fname)
            if path:
                with open(path, "rb") as fh:
                    b64 = base64.b64encode(fh.read()).decode()
                rules.append(f"@font-face{{font-family:'{family}';src:url(data:font/ttf;base64,{b64});"
                             f"font-weight:{weight};font-style:{style};}}")
    if not any("Calibri" in r for r in rules):
        warn("  ! Calibri not found locally: line breaks are measured with the CSS fallback font")
    return "\n".join(rules)


def tag_sections(sections: list):
    """Give every element a stable id so browser measurements can be matched back to the soup."""
    for si, sec in enumerate(sections):
        sec["data-m2p-slide"] = str(si)
        for k, el in enumerate(sec.find_all(True)):
            el["data-m2p"] = f"{si}.{k}"


def measure_layout(chrome: str, css: str, sections: list, html_dir: str) -> dict:
    tmp = tempfile.mkdtemp(prefix="marp2pptx_layout_")
    base = Path(html_dir).as_uri().rstrip("/") + "/"
    body = "".join('<div id=":$p"><svg data-marpit-svg="" viewBox="0 0 1280 720" width="1280" height="720" '
                   'style="display:block"><foreignObject width="1280" height="720">' + str(sec) +
                   "</foreignObject></svg></div>" for sec in sections)
    page = os.path.join(tmp, "layout.html")
    with open(page, "w", encoding="utf-8") as fh:
        fh.write(f"<!DOCTYPE html><html><head><meta charset='utf-8'><base href='{base}'>"
                 f"<style>{css}</style><style>{office_font_css()}</style>"
                 "<style>html,body{margin:0;padding:0;background:#fff}"
                 "[data-marpit-fragment]{visibility:visible!important;opacity:1!important}</style></head><body>"
                 f"{body}<pre id='m2p-out'></pre><script>{MEASURE_JS}</script></body></html>")
    warn(f"Measuring the layout of {len(sections)} slides with headless Chrome ...")
    cmd = [chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
           "--no-default-browser-check", "--allow-file-access-from-files", "--window-size=1280,720",
           f"--user-data-dir={os.path.join(tmp, 'profile')}", "--virtual-time-budget=30000",
           "--timeout=240000", "--dump-dom", Path(page).as_uri()]
    try:
        dom = subprocess.run(cmd, capture_output=True, text=True, timeout=600, creationflags=NO_WINDOW).stdout
        m = re.search(r"<pre id=\"m2p-out\">(.*?)</pre>", dom, re.S)
        data = json.loads(htmllib.unescape(m.group(1))) if m and m.group(1).strip() else {}
    except Exception as exc:
        warn(f"  ! layout measurement failed ({exc})")
        data = {}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if not data:
        warn("  ! no layout measured: falling back to the approximate layout engine")
    return data


def parse_css_color(s: str | None):
    s = (s or "").strip()
    m = re.match(r"rgba?\(([^)]+)\)", s)
    scale = 1.0
    if not m:
        m = re.match(r"color\(srgb\s+([^)]+)\)", s)
        scale = 255.0
    if not m:
        return None
    parts = [p for p in re.split(r"[,\s/]+", m.group(1).strip()) if p]
    try:
        r, g, b = (float(v) * scale for v in parts[:3])
        a = float(parts[3]) if len(parts) > 3 else 1.0
    except ValueError:
        return None
    return r, g, b, a


def css_to_hex(s: str | None, under: str = "FFFFFF", min_alpha: float = 0.02) -> str | None:
    """Computed CSS colour -> opaque hex, alpha-blended over `under`."""
    c = parse_css_color(s)
    if not c or c[3] < min_alpha:
        return None
    r, g, b, a = c
    u = [int(under[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{max(0, min(255, round(v * a + w * (1 - a)))):02X}" for v, w in zip((r, g, b), u))


def ppt_font(css_family: str | None) -> str:
    first = (css_family or "Calibri").split(",")[0].strip().strip("'\"")
    return FONT_ALIASES.get(first.lower(), first)


def render_svg_png(svg: Tag, w: float, h: float, color: str) -> bytes | None:
    if cairosvg is None or w < 1 or h < 1:
        return None
    markup = str(svg)
    for low, camel in (("viewbox", "viewBox"), ("preserveaspectratio", "preserveAspectRatio")):
        markup = markup.replace(f" {low}=", f" {camel}=")
    markup = markup.replace("currentColor", "#" + color)
    markup = re.sub(r"\s(width|height|style)=\"[^\"]*\"", "", markup, count=3)
    markup = markup.replace("<svg", f'<svg width="{w:.1f}px" height="{h:.1f}px"', 1)
    try:
        return cairosvg.svg2png(bytestring=markup.encode(), scale=3)
    except Exception:
        return None


def set_picture_background(slide, png: bytes):
    """Make the PNG the slide's own background (Format Background > Picture), not a shape."""
    _, rid = slide.part.get_or_add_image_part(io.BytesIO(png))
    csld = slide._element.cSld
    old = csld.find(qn("p:bg"))
    if old is not None:
        csld.remove(old)
    csld.insert(0, parse_xml(
        f'<p:bg {nsdecls("p", "a", "r")}><p:bgPr><a:blipFill dpi="0" rotWithShape="1">'
        f'<a:blip r:embed="{rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
        '<a:effectLst/></p:bgPr></p:bg>'))


# --------------------------------------------------------------------------
# HTML section -> blocks
# --------------------------------------------------------------------------
INLINE_STYLE = {"strong": "bold", "b": "bold", "mark": "bold", "em": "italic", "i": "italic", "code": "code"}
BLOCK_TAGS = {"p", "ul", "ol", "table", "pre", "blockquote", "div", "h1", "h2", "h3", "h4", "h5", "h6", "figure", "hr"}


def style_height(tag: Tag) -> float | None:
    m = re.search(r"(?:^|;)\s*(?:max-)?height\s*:\s*([\d.]+)px", tag.get("style", ""))
    if m:
        return float(m.group(1))
    if tag.get("height", "").isdigit():
        return float(tag["height"])
    return None


class SectionParser:
    """Turns the children of a Marp <section> into layout blocks."""

    def __init__(self, base_px: float, pal: dict):
        self.base = base_px
        self.pal = pal

    # ---- inline content -------------------------------------------------
    def inline(self, nodes, style: dict, out: list, extras: list):
        for node in nodes:
            if isinstance(node, Comment):
                continue
            if isinstance(node, NavigableString):
                t = re.sub(r"\s+", " ", str(node))
                if t:
                    out.append(Run(t, **style))
                continue
            name = node.name
            if name == "br":
                out.append(Run(br=True))
            elif name == "img":
                extras.append(self.image_block([node]))
            elif name == "mjx-container":
                extras.append(self.math_block(node))
            elif name == "video":
                extras.append({"kind": "video", "src": node.get("src", "")})
            elif name in ("ul", "ol", "table", "pre", "style", "script"):
                continue
            else:
                s = dict(style)
                if name in INLINE_STYLE:
                    s[INLINE_STYLE[name]] = True
                if name == "a" and node.get("href"):
                    s["href"] = node["href"]
                self.inline(node.children, s, out, extras)

    def para(self, nodes, **kw):
        runs, extras = [], []
        self.inline(nodes, {}, runs, extras)
        runs = clean_runs(runs)
        return (Para(runs, **kw) if runs else None), extras

    # ---- blocks -----------------------------------------------------------
    def heading(self, tag: Tag):
        lvl = int(tag.name[1])
        spec = {
            1: (1.4, HEAD_FONT, self.pal["heading"], False),
            2: (1.15, BODY_FONT, self.pal["h2"], True),
            3: (1.05, BODY_FONT, self.pal["h3"], True),
            4: (0.95, BODY_FONT, None, True),
            5: (0.9, BODY_FONT, None, True),
            6: (0.85, BODY_FONT, None, True),
        }[lvl]
        size, font, color, bold = spec
        return self.para(tag.children, size=self.base * size, font=font, color=color, bold=bold, space_after=0.35)

    def image_block(self, imgs: list) -> dict:
        return {
            "kind": "images",
            "items": [{"src": i.get("src", ""), "title": i.get("title"), "max_h": style_height(i)} for i in imgs],
        }

    def math_block(self, node: Tag) -> dict:
        svg = node.find("svg")
        if svg is None or cairosvg is None:
            return {"kind": "paras", "paras": [Para([Run("[equation]", italic=True)], size=self.base)]}
        markup = str(svg)
        # lxml lower-cases attribute names; restore the SVG camelCase ones
        for low, camel in (("viewbox", "viewBox"), ("preserveaspectratio", "preserveAspectRatio")):
            markup = markup.replace(f" {low}=", f" {camel}=")
        vb = [float(v) for v in (svg.get("viewbox") or svg.get("viewBox") or "0 0 1000 1000").split()]
        w, h = vb[2] / 1000 * self.base, vb[3] / 1000 * self.base       # MathJax: 1000 units = 1em
        markup = markup.replace("currentColor", "#" + self.pal["fg"])
        markup = re.sub(r'\swidth="[^"]*"', f' width="{w:.1f}px"', markup, count=1)
        markup = re.sub(r'\sheight="[^"]*"', f' height="{h:.1f}px"', markup, count=1)
        markup = re.sub(r'\sstyle="[^"]*"', "", markup, count=1)
        try:
            png = cairosvg.svg2png(bytestring=markup.encode(), scale=3)
        except Exception:
            return {"kind": "paras", "paras": [Para([Run("[equation]", italic=True)], size=self.base)]}
        return {"kind": "math", "png": png, "w": w, "h": h}

    def list_paras(self, lst: Tag, level: int, paras: list, extras: list):
        ordered = lst.name == "ol"
        start = int(lst.get("start", 1)) if str(lst.get("start", "1")).isdigit() else 1
        for li in lst.find_all("li", recursive=False):
            runs, nested = [], []
            for child in li.children:
                if isinstance(child, Tag) and child.name in ("ul", "ol"):
                    nested.append(child)
                elif isinstance(child, Tag) and child.name == "p":
                    if runs:
                        runs.append(Run(br=True))
                    self.inline(child.children, {}, runs, extras)
                else:
                    self.inline([child], {}, runs, extras)
            runs = clean_runs(runs)
            if runs:
                bullet = ("ol", level, start) if ordered else ("ul", level)
                paras.append(Para(runs, size=self.base * (1 if level == 0 else 0.9), bullet=bullet, space_after=0.3))
            for sub in nested:
                self.list_paras(sub, level + 1, paras, extras)

    def table_block(self, table: Tag) -> dict:
        rows = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"], recursive=False)
            if not cells:
                continue
            parsed = []
            for c in cells:
                runs, _ = [], []
                self.inline(c.children, {}, runs, _)
                align = "left"
                m = re.search(r"text-align:\s*(left|center|right)", c.get("style", "")) or (
                    re.match(r"(left|center|right)$", c.get("align", "")))
                if m:
                    align = m.group(1)
                parsed.append({"runs": clean_runs(runs), "align": align})
            rows.append({"header": all(c.name == "th" for c in cells), "cells": parsed})
        return {"kind": "table", "rows": rows}

    def parse(self, nodes) -> list:
        blocks = []

        def add_para(p):
            if p is None:
                return
            if blocks and blocks[-1]["kind"] == "paras":
                blocks[-1]["paras"].append(p)
            else:
                blocks.append({"kind": "paras", "paras": [p]})

        for node in nodes:
            if isinstance(node, Comment):
                continue
            if isinstance(node, NavigableString):
                if str(node).strip():
                    add_para(self.para([node], size=self.base)[0])
                continue
            name = node.name
            if name in ("footer", "style", "script", "hr"):
                continue
            if re.fullmatch(r"h[1-6]", name):
                p, extras = self.heading(node)
                add_para(p)
                blocks.extend(extras)
            elif name == "p":
                imgs = node.find_all("img", recursive=False)
                text_left = "".join(str(c) for c in node.children if not (isinstance(c, Tag) and c.name in ("img", "br"))).strip()
                if imgs and not text_left:
                    blocks.append(self.image_block(imgs))
                    continue
                p, extras = self.para(node.children, size=self.base)
                add_para(p)
                blocks.extend(extras)
            elif name in ("ul", "ol"):
                paras, extras = [], []
                self.list_paras(node, 0, paras, extras)
                for p in paras:
                    add_para(p)
                blocks.extend(extras)
            elif name == "table":
                blocks.append(self.table_block(node))
            elif name == "pre":
                blocks.append({"kind": "code", "text": node.get_text().rstrip("\n")})
            elif name == "blockquote":
                inner = self.parse(node.children)
                paras = [p for b in inner if b["kind"] == "paras" for p in b["paras"]]
                if paras:
                    blocks.append({"kind": "quote", "paras": paras})
                blocks.extend(b for b in inner if b["kind"] != "paras")
            elif name == "br":
                if blocks and blocks[-1]["kind"] == "paras":   # vertical spacer
                    add_para(Para([], size=self.base * 0.5, space_after=0))
            elif name == "img":
                blocks.append(self.image_block([node]))
            elif name == "video":
                blocks.append({"kind": "video", "src": node.get("src", "")})
            elif name == "mjx-container":
                blocks.append(self.math_block(node))
            else:   # div, figure, span, center, ... -> walk through
                if any(isinstance(c, Tag) and c.name in BLOCK_TAGS for c in node.children):
                    blocks.extend(self.parse(node.children))
                else:
                    p, extras = self.para(node.children, size=self.base)
                    add_para(p)
                    blocks.extend(extras)
        return blocks


# --------------------------------------------------------------------------
# Text measurement (heuristic, used to pick font scale and box heights)
# --------------------------------------------------------------------------
def bullet_indent(p: Para, s: float) -> float:
    return p.size * s * 1.1 * (p.bullet[1] + 1) if p.bullet else 0


def para_height(p: Para, width: float, s: float) -> float:
    size = p.size * s
    if not p.runs:
        return size * p.line
    cw = CHAR_W.get(p.font, 0.5) * size * (1.06 if p.bold else 1.0)
    avail = max(width - bullet_indent(p, s), 40)
    lines = 0
    for seg in p.segments():
        # word-wrap waste: assume ~8% of each line is lost to wrapping
        lines += max(1, math.ceil(len(seg) * cw / (avail * 0.92)))
    return lines * size * p.line * 1.2 + p.space_after * size


def paras_height(paras: list, width: float, s: float) -> float:
    if not paras:
        return 0
    return sum(para_height(p, width, s) for p in paras) - paras[-1].space_after * paras[-1].size * s


def text_width(text: str, size: float, font: str) -> float:
    return len(text) * CHAR_W.get(font, 0.5) * size


# --------------------------------------------------------------------------
# Slide builder
# --------------------------------------------------------------------------
class DeckBuilder:
    def __init__(self, images: ImageStore, embed_video: bool, themes: dict | None = None,
                 layout: dict | None = None):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = E(CANVAS_W), E(CANVAS_H)
        self.blank = self.prs.slide_layouts[6]
        self.images = images
        self.embed_video = embed_video
        self.credits: list = []
        self.base = 30
        self.themes = themes or {}   # section style key -> {"png": bytes, "pal": computed colours}
        self.layout = layout or {}   # element id -> box and computed style measured by Chrome
        self.forced_breaks = True    # reproduce the browser's line breaks (see --natural-wrap)

    # ---- primitives -----------------------------------------------------
    def textbox(self, slide, paras, x, y, w, h, s, pal, anchor="top", fill=None, pad=0, name=None):
        tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = E(pad)
        tf.margin_top = tf.margin_bottom = E(pad * 0.7 if pad else 0)
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[anchor]
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE   # safety net if the estimate is short
        if fill:
            tb.fill.solid()
            tb.fill.fore_color.rgb = rgb(fill)
        for i, p in enumerate(paras):
            self.write_para(tf.paragraphs[0] if i == 0 else tf.add_paragraph(), p, s, pal)
        return tb

    def write_para(self, para, p: Para, s: float, pal: dict):
        size = p.size * s
        para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[p.align]
        para.line_spacing = p.line
        para.space_after = PT(p.space_after * size)
        if p.bullet:
            self.set_bullet(para, p, s, pal)
        color = p.color or pal["fg"]
        if not p.runs:
            para.font.size = PT(size)
            return
        for r in p.runs:
            if r.br:
                para.add_line_break()
                continue
            run = para.add_run()
            run.text = r.text
            f = run.font
            f.size = PT(size * (0.88 if r.code else 1))
            f.name = CODE_FONT if r.code else p.font
            f.bold = True if (p.bold or r.bold) else None
            f.italic = True if r.italic else None
            f.color.rgb = rgb(color)
            if r.href and re.match(r"(https?|mailto|file):", r.href):
                run.hyperlink.address = r.href

    @staticmethod
    def set_bullet(para, p: Para, s: float, pal: dict):
        size = p.size * s
        level = p.bullet[1]
        pPr = para._p.get_or_add_pPr()
        step = size * 1.1
        pPr.set("marL", str(int(E(step * (level + 1)))))
        pPr.set("indent", str(int(-E(step))))
        clr = etree.SubElement(pPr, qn("a:buClr"))
        etree.SubElement(clr, qn("a:srgbClr")).set("val", pal["accent"])
        if p.bullet[0] == "ol":
            etree.SubElement(pPr, qn("a:buFont")).set("typeface", "+mj-lt")
            num = etree.SubElement(pPr, qn("a:buAutoNum"))
            num.set("type", "arabicPeriod")
            num.set("startAt", str(p.bullet[2]))
        else:
            etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
            etree.SubElement(pPr, qn("a:buChar")).set("char", "•" if level % 2 == 0 else "–")

    def picture(self, slide, item, x, y, w, h, pal, mode="contain", caption=True):
        """Place one image inside the box. Returns the drawn rectangle (x, y, w, h)."""
        img = self.images.get(item["src"])
        cap_h = 18 if (caption and item.get("title") and mode == "contain") else 0
        if img is None:
            shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(y), E(w), E(max(h, 40)))
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb(pal["card_bg"])
            shp.line.color.rgb = rgb(pal["border"])
            tf = shp.text_frame
            tf.word_wrap = True
            tf.text = "Image not found: " + urllib.parse.unquote(os.path.basename(item["src"]))[:80]
            tf.paragraphs[0].runs[0].font.size = Pt(11)
            tf.paragraphs[0].runs[0].font.color.rgb = rgb(text_on(pal["card_bg"]))
            return x, y, w, h
        data, iw, ih = img
        if mode == "cover":
            pic = slide.shapes.add_picture(io.BytesIO(data), E(x), E(y), E(w), E(h))
            scale = max(w / iw, h / ih)
            crop_x = max(0.0, (1 - w / (iw * scale)) / 2)
            crop_y = max(0.0, (1 - h / (ih * scale)) / 2)
            pic.crop_left = pic.crop_right = crop_x
            pic.crop_top = pic.crop_bottom = crop_y
            dx, dy, dw, dh = x, y, w, h
        else:
            bh = h - cap_h
            scale = min(w / iw, bh / ih)
            dw, dh = iw * scale, ih * scale
            dx, dy = x + (w - dw) / 2, y + (bh - dh) / 2
            pic = slide.shapes.add_picture(io.BytesIO(data), E(dx), E(dy), E(dw), E(dh))
            if cap_h:
                cap = Para([Run(item["title"])], size=13, color=SU["cool_grey"] if luminance(pal["bg"]) > 0.5 else "D9D9D9",
                           align="right", space_after=0, line=1.0)
                self.textbox(slide, [cap], dx, dy + dh + 2, max(dw, 260), 16, 1, pal)
        pic.name = "Picture " + os.path.basename(urllib.parse.unquote(item["src"]))[:40]
        if item.get("title"):
            self.credits.append(item["title"])
        return dx, dy, dw, dh

    def video(self, slide, block, x, y, w, h, pal):
        src = urllib.parse.unquote(block["src"].replace("file://", ""))
        vw = min(w, h * 16 / 9)
        vh = vw * 9 / 16
        vx, vy = x + (w - vw) / 2, y + (h - vh) / 2
        if self.embed_video and os.path.isfile(src):
            mime = "video/quicktime" if src.lower().endswith(".mov") else "video/mp4"
            slide.shapes.add_movie(src, E(vx), E(vy), E(vw), E(vh), mime_type=mime)
            return
        p = Para([Run("▶ Video: " + os.path.basename(src), href="file://" + urllib.parse.quote(src))],
                 size=24, color=text_on(pal["card_bg"]), align="center")
        self.textbox(slide, [p], vx, vy, vw, vh, 1, pal, anchor="middle", fill=pal["card_bg"], pad=12)

    # ---- tables & code ----------------------------------------------------
    def table(self, slide, block, x, y, w, s, pal, draw=True) -> float:
        rows = block["rows"]
        ncols = max(len(r["cells"]) for r in rows)
        size = max(self.base * s * 0.62, 11)
        pad = 6
        weights = []
        for c in range(ncols):
            longest = max((len(runs_text(r["cells"][c]["runs"])) for r in rows if c < len(r["cells"])), default=4)
            weights.append(min(max(longest, 4), 45))
        col_w = [w * wt / sum(weights) for wt in weights]
        heights = []
        for r in rows:
            hmax = 0
            for c, cell in enumerate(r["cells"]):
                txt = runs_text(cell["runs"]) or " "
                lines = sum(max(1, math.ceil(len(seg) * CHAR_W[BODY_FONT] * size * (1.06 if r["header"] else 1) /
                                             ((col_w[c] - 2 * pad) * 0.92))) for seg in txt.split("\n"))
                hmax = max(hmax, lines * size * 1.2 + 2 * pad)
            heights.append(hmax)
        total = sum(heights)
        if not draw:
            return total
        shape = slide.shapes.add_table(len(rows), ncols, E(x), E(y), E(w), E(total))
        tbl = shape.table
        tbl.first_row = rows[0]["header"]
        tbl.horz_banding = False
        for c in range(ncols):
            tbl.columns[c].width = E(col_w[c])
        body_i = 0
        for ri, r in enumerate(rows):
            tbl.rows[ri].height = E(heights[ri])
            if r["header"]:
                fill = pal["accent"]
            else:
                fill = pal["card_bg"] if body_i % 2 == 0 else pal["bg"]
                body_i += 1
            fg = text_on(fill)
            for c in range(ncols):
                cell = tbl.cell(ri, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(fill)
                cell.margin_left = cell.margin_right = E(pad)
                cell.margin_top = cell.margin_bottom = E(pad * 0.6)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cdata = r["cells"][c] if c < len(r["cells"]) else {"runs": [], "align": "left"}
                p = Para(cdata["runs"], size=size, color=fg, bold=r["header"], align=cdata["align"], space_after=0, line=1.0)
                tf = cell.text_frame
                tf.word_wrap = True
                self.write_para(tf.paragraphs[0], p, 1, pal)
        return total

    def code(self, slide, block, x, y, w, s, pal, draw=True) -> float:
        size = max(self.base * s * 0.6, 10)
        pad = 14
        lines = block["text"].split("\n")
        n = sum(max(1, math.ceil(len(l) * CHAR_W[CODE_FONT] * size / (w - 2 * pad))) for l in lines)
        total = n * size * 1.18 + 2 * pad
        if draw:
            paras = [Para([Run(l if l else " ")], size=size, font=CODE_FONT, color=text_on(pal["card_bg"]),
                          space_after=0, line=1.0) for l in lines]
            self.textbox(slide, paras, x, y, w, total, 1, pal, fill=pal["card_bg"], pad=pad, name="Code")
        return total

    # ---- flow layout ------------------------------------------------------
    def flow(self, slide, blocks, x, y, w, h, base, pal, fill_images=False, valign="top"):
        """Stack blocks vertically in the box, shrinking text until everything fits."""
        self.base = base
        if not blocks:
            return
        flex_kinds = ("images", "video")
        QP = 18

        def fixed_h(b, s):
            k = b["kind"]
            if k == "paras":
                return paras_height(b["paras"], w, s)
            if k == "quote":
                return paras_height(b["paras"], w - 2 * QP, s) + 2 * QP * 0.7 + 8
            if k == "code":
                return self.code(slide, b, x, y, w, s, pal, draw=False)
            if k == "table":
                return self.table(slide, b, x, y, w, s, pal, draw=False)
            if k == "math":
                return min(b["h"] * s, b["h"] * w / b["w"]) + 8
            return 0

        def row_aspects(b):
            out = []
            for it in b["items"]:
                img = self.images.get(it["src"])
                out.append((img[1] / img[2], img[2]) if img else (16 / 9, 300))
            return out

        def pref_h(b):
            if b["kind"] == "video":
                return w * 9 / 16
            asp = row_aspects(b)
            cap = 18 if any(it.get("title") for it in b["items"]) else 0
            if fill_images:
                return 10 ** 6
            nat = max(it["max_h"] or ih for it, (_, ih) in zip(b["items"], asp))
            fit_w = (w - GAP * (len(asp) - 1)) / sum(a for a, _ in asp)
            return min(nat, fit_w) + cap

        flex = [b for b in blocks if b["kind"] in flex_kinds]
        n_gaps = len(blocks) - 1
        for s in SCALES:
            fixed = sum(fixed_h(b, s) for b in blocks if b["kind"] not in flex_kinds)
            gaps = GAP * 0.7 * s * n_gaps
            min_flex = sum(min(pref_h(b), 150 if len(flex) < 3 else 90) for b in flex)
            if fixed + gaps + min_flex <= h:
                break
        remaining = max(h - fixed - gaps, 60 * len(flex))
        prefs = [pref_h(b) for b in flex]
        if sum(prefs) <= remaining:
            alloc = prefs
        else:
            total = sum(min(p, remaining) for p in prefs)
            alloc = [remaining * min(p, remaining) / total for p in prefs]
        heights, fi = [], 0
        for b in blocks:
            if b["kind"] in flex_kinds:
                heights.append(alloc[fi])
                fi += 1
            else:
                heights.append(fixed_h(b, s))
        used = sum(heights) + gaps
        cy = y + max(0, (h - used) / 2) if valign == "middle" else y
        for b, bh in zip(blocks, heights):
            k = b["kind"]
            if k == "paras":
                self.textbox(slide, b["paras"], x, cy, w, bh + 4, s, pal)
            elif k == "quote":
                self.textbox(slide, b["paras"], x, cy, w, bh, s, pal, fill=pal["card_bg"], pad=QP)
            elif k == "code":
                self.code(slide, b, x, cy, w, s, pal)
            elif k == "table":
                self.table(slide, b, x, cy, w, s, pal)
            elif k == "math":
                mw, mh = min(b["w"] * s, w), bh - 8
                mw = min(mw, mh * b["w"] / b["h"])
                slide.shapes.add_picture(io.BytesIO(b["png"]), E(x + (w - mw) / 2), E(cy + 4), E(mw), E(mw * b["h"] / b["w"]))
            elif k == "video":
                self.video(slide, b, x, cy, w, bh, pal)
            elif k == "images":
                self.image_row(slide, b, x, cy, w, bh, pal)
            cy += bh + GAP * 0.7 * s

    def image_row(self, slide, b, x, y, w, h, pal):
        items = b["items"]
        asp = []
        for it in items:
            img = self.images.get(it["src"])
            asp.append(img[1] / img[2] if img else 16 / 9)
        cap = 18 if any(it.get("title") for it in items) else 0
        g = 16
        rh = min(h - cap, (w - g * (len(items) - 1)) / sum(asp))
        widths = [rh * a for a in asp]
        cx = x + (w - sum(widths) - g * (len(items) - 1)) / 2
        for it, iw in zip(items, widths):
            self.picture(slide, it, cx, y + (h - cap - rh) / 2, iw, rh + cap, pal)
            cx += iw + g

    # ---- titles -----------------------------------------------------------
    def fitted_title(self, runs, size, font, width, max_lines=2, min_scale=0.55):
        text = runs_text(runs)
        s = 1.0
        while s > min_scale:
            lines = sum(max(1, math.ceil(text_width(seg, size * s, font) / (width * 0.95))) for seg in text.split("\n"))
            if lines <= max_lines:
                return size * s, lines
            s -= 0.05
        return size * min_scale, max_lines

    def title(self, slide, runs, base, pal, x=PAD_X, y=PAD_T, w=CONTENT_W, align="left") -> float:
        size, lines = self.fitted_title(runs, base * 1.7, HEAD_FONT, w)
        h = lines * size * 1.15 * 1.1
        p = Para(runs, size=size, font=HEAD_FONT, color=pal["heading"], space_after=0, line=1.0, align=align)
        self.textbox(slide, [p], x, y, w, h, 1, pal, anchor="bottom" if lines == 1 else "top", name="Title")
        return y + h

    # ---- slide types --------------------------------------------------------
    def build(self, section: Tag, note: str):
        classes = section.get("class", [])
        if isinstance(classes, str):
            classes = classes.split()
        color_n = next((int(c.split("-")[1]) for c in classes if re.fullmatch(r"color-\d", c)), 1)
        pal = dict(PALETTES.get(color_n, PALETTES[1]))
        theme = self.themes.get(section_key(section))
        if theme:
            pal.update(theme["pal"])          # colours computed by the browser for this style
        pal.setdefault("h2", pal["heading"])
        pal.setdefault("h3", pal["heading"])
        base = 21 if "small" in classes else 25 if "medium" in classes else 30
        slide = self.prs.slides.add_slide(self.blank)
        has_bg = bool(theme and theme["png"])
        if has_bg:
            set_picture_background(slide, theme["png"])
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = rgb(pal["bg"])
        self.credits = []

        if "slide" + section.get("data-m2p-slide", "x") in self.layout:
            self.build_exact(slide, section, pal)
            self.notes(slide, note)
            return

        footer = section.find("footer", recursive=False)
        footer_text = footer.get_text(" ", strip=True) if footer else ""
        nodes = [c for c in section.children if not (isinstance(c, Tag) and c.name == "footer")]
        parser = SectionParser(base, pal)

        full_bleed = "image-full" in classes and "bg" in classes
        if "title" in classes:
            self.layout_title(slide, nodes, parser, base, pal)
        elif "chapter" in classes:
            self.layout_chapter(slide, nodes, parser, base, pal)
        elif "callout" in classes:
            self.layout_callout(slide, nodes, parser, base, pal, classes)
        else:
            self.layout_content(slide, nodes, parser, base, pal, classes)

        if not full_bleed:
            self.footer(slide, footer_text, pal, draw_rule=not has_bg)
        self.notes(slide, note)

    def notes(self, slide, note):
        notes = note.strip()
        if self.credits:
            notes = (notes + "\n\n" if notes else "") + "Image credits: " + "; ".join(dict.fromkeys(self.credits))
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    def pull(self, nodes, tag_name):
        """Remove and return the first direct child <tag_name>."""
        for i, n in enumerate(nodes):
            if isinstance(n, Tag) and n.name == tag_name:
                return n, nodes[:i] + nodes[i + 1:]
        return None, nodes

    def layout_title(self, slide, nodes, parser, base, pal):
        """Theme CSS: padding-left 15%; h1 2.7em serif; h2 and p indented 26%; block centred vertically."""
        x0 = CANVAS_W * 0.15
        avail = CANVAS_W - x0 - PAD_X
        x_ind = x0 + avail * 0.26
        items, pending_gap = [], 0.0          # (paras, x, w, gap_before)
        tags = [n for n in nodes if isinstance(n, Tag)]
        last_p = next((t for t in reversed(tags) if t.name == "p"), None)
        for t in tags:
            if t.name == "br":
                pending_gap += base * 1.4
                continue
            n_before = len(items)
            if t.name == "h1":
                p, _ = parser.para(t.children, size=base * 2.7, font=HEAD_FONT, color=pal["heading"], space_after=0, line=1.12)
                if p:
                    p.size, _ = self.fitted_title(p.runs, base * 2.7, HEAD_FONT, avail, max_lines=3)
                    items.append(([p], x0, avail, pending_gap))
            elif t.name == "h2":
                p, _ = parser.para(t.children, size=base, color=pal["h2"], space_after=0, line=1.4)
                if p:
                    items.append(([p], x_ind, CANVAS_W - x_ind - PAD_X, pending_gap))
            elif re.fullmatch(r"h[3-6]", t.name):
                p, _ = parser.para(t.children, size=base, bold=True, color=pal["h3"] if t.name == "h3" else None,
                                   space_after=0)
                if p:
                    items.append(([p], x0, avail, pending_gap))
            elif t.name == "p":
                p, _ = parser.para(t.children, size=base * 0.9, space_after=0)
                if p:
                    if t is last_p:
                        for r in p.runs:
                            r.italic = True
                    items.append(([p], x_ind, CANVAS_W - x_ind - PAD_X, pending_gap))
            else:
                paras = [p for b in parser.parse([t]) if b["kind"] == "paras" for p in b["paras"]]
                if paras:
                    items.append((paras, x0, avail, pending_gap))
            if len(items) > n_before:
                pending_gap = 0.0
        if not items:
            return
        heights = [paras_height(paras, w, 1) * 1.05 for paras, _, w, _ in items]
        total = sum(heights) + sum(g for *_, g in items) + GAP * (len(items) - 1)
        y = (PAD_T + BODY_BOTTOM) / 2 - total / 2
        for (paras, x, w, gap), h in zip(items, heights):
            y += gap
            self.textbox(slide, paras, x, y, w, h, 1, pal, name="Title" if paras[0].font == HEAD_FONT else None)
            y += h + GAP

    def layout_chapter(self, slide, nodes, parser, base, pal):
        h1, nodes = self.pull(nodes, "h1")
        h2, nodes = self.pull(nodes, "h2")
        rest = parser.parse(nodes)
        x0 = CANVAS_W * 0.15
        avail_w = CANVAS_W - x0 - PAD_X
        h1_runs = clean_runs(self._runs(parser, h1)) if h1 else []
        h2_runs = clean_runs(self._runs(parser, h2)) if h2 else []
        if h1_runs and h2_runs:
            size1 = base * 3
            w1 = min(text_width(runs_text(h1_runs), size1, HEAD_FONT) * 1.1 + 10, avail_w * 0.4)
            size1, l1 = self.fitted_title(h1_runs, size1, HEAD_FONT, w1, max_lines=2)
            x2 = x0 + w1 + 48
            size2, l2 = self.fitted_title(h2_runs, base * 1.8, HEAD_FONT, CANVAS_W - x2 - PAD_X, max_lines=3)
            row_h = max(l1 * size1, l2 * size2) * 1.25
        else:
            runs = h1_runs or h2_runs
            size1, l1 = self.fitted_title(runs, base * 2.2, HEAD_FONT, avail_w, max_lines=3) if runs else (0, 0)
            row_h = l1 * size1 * 1.25
        rest_h = min(sum(paras_height(b["paras"], avail_w, 1) for b in rest if b["kind"] == "paras") + (300 if any(
            b["kind"] != "paras" for b in rest) else 0), BODY_BOTTOM - PAD_T - row_h - GAP)
        total = row_h + (GAP + rest_h if rest else 0)
        y0 = max(PAD_T, (CANVAS_H - total) / 2)
        if h1_runs and h2_runs:
            self.textbox(slide, [Para(h1_runs, size=size1, font=HEAD_FONT, color=pal["accent"], space_after=0, line=1.0)],
                         x0, y0, w1, row_h, 1, pal, anchor="middle", name="Title")
            self.textbox(slide, [Para(h2_runs, size=size2, font=HEAD_FONT, color=pal["heading"], space_after=0, line=1.05)],
                         x2, y0, CANVAS_W - x2 - PAD_X, row_h, 1, pal, anchor="middle", name="Subtitle")
        elif h1_runs or h2_runs:
            self.textbox(slide, [Para(h1_runs or h2_runs, size=size1, font=HEAD_FONT, color=pal["accent"], space_after=0, line=1.05)],
                         x0, y0, avail_w, row_h, 1, pal, anchor="middle", name="Title")
        if rest:
            self.flow(slide, rest, x0, y0 + row_h + GAP, avail_w, BODY_BOTTOM - (y0 + row_h + GAP), base, pal)

    def layout_callout(self, slide, nodes, parser, base, pal, classes):
        h1, nodes = self.pull(nodes, "h1")
        size_cls = next((int(c.rsplit("-", 1)[1]) for c in classes if re.fullmatch(r"callout-size-\d+", c)), None)
        rest = parser.parse(nodes)
        for b in rest:
            for p in b.get("paras", []):
                p.align = "center"
        runs = clean_runs(self._runs(parser, h1)) if h1 else []
        rest_h = sum(paras_height(b["paras"], CONTENT_W, 1) for b in rest if b["kind"] == "paras")
        rest_h += sum(250 for b in rest if b["kind"] != "paras")
        if runs:
            size, lines = self.fitted_title(runs, size_cls or base * 3.6, HEAD_FONT, CONTENT_W, max_lines=3, min_scale=0.3)
            th = lines * size * 1.15
            total = th + (GAP + rest_h if rest else 0)
            y0 = max(PAD_T, (CANVAS_H - 40 - total) / 2)
            self.textbox(slide, [Para(runs, size=size, font=HEAD_FONT, color=pal["heading"], align="center", space_after=0, line=1.05)],
                         PAD_X, y0, CONTENT_W, th, 1, pal, anchor="middle", name="Title")
            if rest:
                self.flow(slide, rest, PAD_X, y0 + th + GAP, CONTENT_W, BODY_BOTTOM - (y0 + th + GAP), base, pal)
        elif rest:
            self.flow(slide, rest, PAD_X, PAD_T, CONTENT_W, BODY_BOTTOM - PAD_T, base, pal, valign="middle")

    def layout_content(self, slide, nodes, parser, base, pal, classes):
        h1, nodes = self.pull(nodes, "h1")
        blocks = parser.parse(nodes)
        title_runs = clean_runs(self._runs(parser, h1)) if h1 else []

        if "image-full" in classes:
            imgs = [it for b in blocks if b["kind"] == "images" for it in b["items"]]
            others = [b for b in blocks if b["kind"] != "images"]
            if imgs:
                if "bg" in classes:
                    self.picture(slide, imgs[0], 0, 0, CANVAS_W, CANVAS_H, pal, mode="cover")
                else:
                    self.picture(slide, imgs[0], 0, 0, CANVAS_W, CANVAS_H - 64, pal, caption=False)
            if title_runs:
                self.title(slide, title_runs, base, pal)
            if others:
                self.flow(slide, others, PAD_X, BODY_BOTTOM - 200, CONTENT_W, 200, base, pal)
            return

        top = self.title(slide, title_runs, base, pal) + GAP if title_runs else PAD_T
        box_h = BODY_BOTTOM - top

        split = next((re.fullmatch(r"(image-text|text-image)-(\d+)", c) for c in classes
                      if re.fullmatch(r"(image-text|text-image)-(\d+)", c)), None)
        gallery = "gallery" in classes
        if split:
            frac = int(split.group(2)) / 100
            img_blocks = [b for b in blocks if b["kind"] in ("images", "video")]
            txt_blocks = [b for b in blocks if b["kind"] not in ("images", "video")]
            if img_blocks and txt_blocks:
                img_w = (CONTENT_W - GAP) * frac
                txt_w = CONTENT_W - GAP - img_w
                img_x, txt_x = (PAD_X, PAD_X + img_w + GAP) if split.group(1) == "image-text" else (PAD_X + txt_w + GAP, PAD_X)
                self.flow(slide, img_blocks, img_x, top, img_w, box_h, base, pal, fill_images=True, valign="middle")
                self.flow(slide, txt_blocks, txt_x, top, txt_w, box_h, base, pal, valign="middle")
                return
        cols_m = next((re.fullmatch(r"cols-(\d+)", c) for c in classes if re.fullmatch(r"cols-(\d+)", c)), None)
        if gallery or cols_m:
            # CSS grid: every direct child element is one cell
            cols = int(cols_m.group(1)) if cols_m else 3
            cell_nodes = [n for n in nodes if isinstance(n, Tag) or (not isinstance(n, Comment) and str(n).strip())]
            legend = []
            if ("with-legend" in classes and cell_nodes and isinstance(cell_nodes[-1], Tag)
                    and cell_nodes[-1].name in ("blockquote", "p") and not cell_nodes[-1].find("img")):
                legend = [{"kind": "paras", "paras": b["paras"]} if b["kind"] == "quote" else b
                          for b in parser.parse([cell_nodes.pop()])]
            cells = [c for c in (parser.parse([n]) for n in cell_nodes) if c]
            legend_h = sum(paras_height(b["paras"], CONTENT_W, 1) for b in legend if b["kind"] == "paras")
            grid_h = box_h - (legend_h + GAP if legend else 0)
            cols = max(1, min(cols, len(cells)))
            nrows = max(1, math.ceil(len(cells) / cols))
            cw = (CONTENT_W - GAP * (cols - 1)) / cols
            ch = (grid_h - GAP * (nrows - 1)) / nrows
            for i, cell in enumerate(cells):
                r, c = divmod(i, cols)
                visual = all(b["kind"] in ("images", "video") for b in cell)
                self.flow(slide, cell, PAD_X + c * (cw + GAP), top + r * (ch + GAP), cw, ch, base, pal,
                          fill_images=True, valign="middle" if visual else "top")
            if legend:
                self.flow(slide, legend, PAD_X, BODY_BOTTOM - legend_h, CONTENT_W, legend_h, base, pal)
            return

        image_fit = "image-fit" in classes
        only_images = all(b["kind"] in ("images", "video") for b in blocks)
        self.flow(slide, blocks, PAD_X, top, CONTENT_W, box_h, base, pal,
                  fill_images=image_fit, valign="middle" if (image_fit or only_images) else "top")

    @staticmethod
    def _runs(parser, tag):
        runs, extras = [], []
        if tag is not None:
            parser.inline(tag.children, {}, runs, extras)
        return runs

    # ======================================================================
    # Exact layout: every element is placed at the box Chrome computed for it
    # ======================================================================
    def breaks(self, runs, rec):
        return apply_breaks(runs, rec.get("lb")) if self.forced_breaks else runs

    def L(self, tag) -> dict | None:
        return self.layout.get(tag.get("data-m2p")) if isinstance(tag, Tag) else None

    @staticmethod
    def scale_of(rec: dict) -> float:
        s = rec.get("s") or 1
        return 1.0 if abs(s - 1) < 0.03 else s

    @staticmethod
    def line_px(rec: dict, scale: float) -> float:
        lh = rec.get("lh", "normal")
        px = float(lh[:-2]) if lh.endswith("px") else rec["fs"] * 1.2
        return px * scale

    @staticmethod
    def text_top(rec: dict, lh: float) -> float | None:
        """Top of the first line box: first glyph run minus CSS half-leading."""
        ft = rec.get("ft")
        return ft[0] - (lh - ft[1]) / 2 if ft else None

    @staticmethod
    def content_box(rec: dict):
        (pt, pr, pb, pl), (bt, br_, bb, bl) = rec["p"], rec["b"]
        return (rec["x"] + bl + pl, rec["y"] + bt + pt,
                max(rec["w"] - bl - pl - br_ - pr, 1), max(rec["h"] - bt - pt - bb - pb, 1))

    def run_style(self, rec, scale, under, href):
        return dict(bold=str(rec.get("fw", "400")).isdigit() and int(rec["fw"]) >= 600 or rec.get("fw") == "bold",
                    italic=rec.get("fi") in ("italic", "oblique"), href=href,
                    color=css_to_hex(rec.get("c"), under), size=rec["fs"] * scale, font=ppt_font(rec.get("ff")),
                    underline="underline" in (rec.get("td") or ""),
                    spacing=float(rec["lsp"][:-2]) * scale if str(rec.get("lsp", "")).endswith("px") else None)

    def exact_runs(self, nodes, rec, scale, under, out, pre=False, href=None):
        """Inline content -> runs, each carrying the computed style of its element."""
        for child in nodes:
            if isinstance(child, Comment):
                continue
            if isinstance(child, NavigableString):
                t = str(child)
                if rec.get("tt") == "uppercase":
                    t = t.upper()
                elif rec.get("tt") == "lowercase":
                    t = t.lower()
                if pre:
                    for i, line in enumerate(t.split("\n")):
                        if i:
                            out.append(Run(br=True))
                        if line:
                            out.append(Run(line, **self.run_style(rec, scale, under, href)))
                else:
                    t = re.sub(r"\s+", " ", t)
                    if t:
                        out.append(Run(t, **self.run_style(rec, scale, under, href)))
                continue
            name = child.name
            if name == "br":
                out.append(Run(br=True))
                continue
            if name in ("img", "video", "mjx-container", "ul", "ol", "table", "pre", "style", "script"):
                continue
            crec = self.L(child)
            if crec is None and child.get("data-m2p"):
                continue            # not rendered (display:none)
            cscale = scale
            if crec and crec.get("d") != "inline" and self.scale_of(crec) != 1:
                cscale = self.scale_of(crec)    # e.g. Marp auto-fit content scaled by an SVG wrapper
            self.exact_runs(child.children, crec or rec, cscale, under, out, pre,
                            child.get("href") if name == "a" else href)

    def exact_textbox(self, slide, paras, x, y, w, h, name=None):
        """paras: dicts with runs, align, lh (px), size (px), lines, bullet, and either `baseline`
        (absolute px, where Chrome drew the first baseline) or `sb` (space before, px).

        PowerPoint snaps spacing to whole points. Percentage line spacing keeps the average
        exact (no drift inside a paragraph); every paragraph is then pinned to the browser's
        baseline, the rounding of its space-before being absorbed paragraph by paragraph."""
        first = paras[0]
        # Line breaks are copied from the browser, so extra width can never re-wrap a line; it only
        # stops PowerPoint from wrapping earlier than Chrome did because of sub-pixel differences.
        slack = max(4.0, w * 0.04) if self.forced_breaks else max(1.5, w * 0.003)
        if first["align"] == "center":
            x -= slack / 2
        elif first["align"] == "right":
            x -= slack
        bottom = y + h
        if first.get("baseline") is not None:
            y = first["baseline"] - ppt_baseline(*para_font(first))
        tb = slide.shapes.add_textbox(E(x), E(y), E(w + slack), E(max(bottom - y, first["lh"])))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        got = 0.0                             # where PowerPoint puts the next paragraph (pt from box top)
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "end": PP_ALIGN.RIGHT,
                              "justify": PP_ALIGN.JUSTIFY}.get(p["align"], PP_ALIGN.LEFT)
            font, size, lh = para_font(p)
            para.line_spacing = ppt_pct(size, lh)                  # percentage spacing
            if i:
                if p.get("baseline") is not None:
                    want = (p["baseline"] - ppt_baseline(font, size, lh) - y) * 0.75
                else:
                    want = got + p.get("sb", 0) * 0.75
                sb_pt = max(0, round(want - got))
                got += sb_pt
                para.space_before = Pt(sb_pt)
            para.space_after = Pt(0)
            got += p.get("lines", 1) * lh * 0.75
            if p.get("bullet"):
                self.exact_bullet(para, **p["bullet"])
            for r in p["runs"]:
                if r.br:
                    para.add_line_break()
                    continue
                run = para.add_run()
                run.text = r.text
                f = run.font
                f.size = PT(r.size or p["size"])
                f.name = r.font or BODY_FONT
                f.bold = True if r.bold else None
                f.italic = True if r.italic else None
                f.underline = True if r.underline else None
                if r.color:
                    f.color.rgb = rgb(r.color)
                rPr = run._r.get_or_add_rPr()
                rPr.set("kern", "100")             # kern like the browser does
                if r.spacing:
                    rPr.set("spc", str(int(round(r.spacing * 75))))   # 1/100 pt
                if r.href and re.match(r"(https?|mailto|file):", r.href):
                    run.hyperlink.address = r.href
        return tb

    @staticmethod
    def exact_bullet(para, marL, indent, style, color, start=1):
        pPr = para._p.get_or_add_pPr()
        pPr.set("marL", str(int(E(marL))))
        pPr.set("indent", str(int(E(indent))))
        if style == "none":
            etree.SubElement(pPr, qn("a:buNone"))
            return
        if color:
            etree.SubElement(etree.SubElement(pPr, qn("a:buClr")), qn("a:srgbClr")).set("val", color)
        auto = {"decimal": "arabicPeriod", "decimal-leading-zero": "arabicPeriod", "lower-alpha": "alphaLcPeriod",
                "lower-latin": "alphaLcPeriod", "upper-alpha": "alphaUcPeriod", "upper-latin": "alphaUcPeriod",
                "lower-roman": "romanLcPeriod", "upper-roman": "romanUcPeriod"}.get(style)
        if auto:
            etree.SubElement(pPr, qn("a:buFont")).set("typeface", "+mj-lt")
            n = etree.SubElement(pPr, qn("a:buAutoNum"))
            n.set("type", auto)
            n.set("startAt", str(start))
        else:
            etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
            etree.SubElement(pPr, qn("a:buChar")).set("char", {"circle": "◦", "square": "▪"}.get(style, "•"))

    def exact_text(self, slide, node, rec, under, scale=None, nodes=None):
        scale = scale or self.scale_of(rec)
        runs = []
        self.exact_runs(node.children if nodes is None else nodes, rec, scale, under, runs)
        runs = self.breaks(clean_runs(runs), rec)
        if not runs:
            return
        x, y, w, h = self.content_box(rec)
        lh = self.line_px(rec, scale)
        top = self.text_top(rec, lh)
        if top is not None:                # e.g. text pushed down inside a stretched grid/flex box
            h, y = y + h - top, top
        para = {"runs": runs, "align": rec.get("ta", "left"), "lh": lh, "size": rec["fs"] * scale}
        para["baseline"] = css_baseline(rec, para_font(para)[0])
        self.exact_textbox(slide, [para], x, y, w, h, name=node.name)

    def decor(self, slide, node, rec, under):
        """Backgrounds and borders drawn by CSS on block elements."""
        bg = css_to_hex(rec.get("bg"), under)
        if bg and rec["w"] > 1 and rec["h"] > 1:
            kind = MSO_SHAPE.ROUNDED_RECTANGLE if rec.get("br") else MSO_SHAPE.RECTANGLE
            shp = slide.shapes.add_shape(kind, E(rec["x"]), E(rec["y"]), E(rec["w"]), E(rec["h"]))
            if rec.get("br"):
                shp.adjustments[0] = min(0.5, rec["br"] / min(rec["w"], rec["h"]))
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb(bg)
            shp.line.fill.background()
            shp.shadow.inherit = False
            shp.name = f"{node.name} background"
        if rec.get("bc"):
            t, r, b, l = rec["b"]
            x, y, w, h = rec["x"], rec["y"], rec["w"], rec["h"]
            for width, colour, box in ((t, rec["bc"][0], (x, y, w, t)), (r, rec["bc"][1], (x + w - r, y, r, h)),
                                       (b, rec["bc"][2], (x, y + h - b, w, b)), (l, rec["bc"][3], (x, y, l, h))):
                hexc = css_to_hex(colour, under)
                if width >= 0.5 and hexc:
                    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *(E(v) for v in box))
                    shp.fill.solid()
                    shp.fill.fore_color.rgb = rgb(hexc)
                    shp.line.fill.background()
                    shp.shadow.inherit = False
                    shp.name = f"{node.name} border"

    def exact_image(self, slide, node, rec, under):
        x, y, w, h = self.content_box(rec)
        src = node.get("src", "")
        img = self.images.get(src)
        if img is None:
            self.picture(slide, {"src": src}, x, y, w, h, {"card_bg": "F4F4F4", "border": "B6B1A9", "bg": under},
                         caption=False)
            return
        data, iw, ih = img
        fit = rec.get("of", "fill")
        if fit in ("contain", "scale-down", "none"):
            s = min(w / iw, h / ih)
            dw, dh = iw * s, ih * s
            pic = slide.shapes.add_picture(io.BytesIO(data), E(x + (w - dw) / 2), E(y + (h - dh) / 2), E(dw), E(dh))
        elif fit == "cover":
            pic = slide.shapes.add_picture(io.BytesIO(data), E(x), E(y), E(w), E(h))
            s = max(w / iw, h / ih)
            pic.crop_left = pic.crop_right = max(0.0, (1 - w / (iw * s)) / 2)
            pic.crop_top = pic.crop_bottom = max(0.0, (1 - h / (ih * s)) / 2)
        else:
            pic = slide.shapes.add_picture(io.BytesIO(data), E(x), E(y), E(w), E(h))
        pic.name = "Picture " + os.path.basename(urllib.parse.unquote(src))[:40]
        self.clip_to_slide(pic)
        if node.get("title"):
            self.credits.append(node["title"])

    @staticmethod
    def clip_to_slide(pic):
        """The slide clips overflowing content (overflow: hidden); crop the picture the same way."""
        x, y, w, h = pic.left, pic.top, pic.width, pic.height
        W, H = E(CANVAS_W), E(CANVAS_H)
        if w <= 0 or h <= 0 or (x >= 0 and y >= 0 and x + w <= W and y + h <= H):
            return
        cl, ct, cr, cb = pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom
        sx, sy = (1 - cl - cr) / w, (1 - ct - cb) / h      # crop fraction per EMU of shape
        if x < 0:
            cl += -x * sx
        if y < 0:
            ct += -y * sy
        if x + w > W:
            cr += (x + w - W) * sx
        if y + h > H:
            cb += (y + h - H) * sy
        nx, ny = max(x, 0), max(y, 0)
        nw, nh = min(x + w, W) - nx, min(y + h, H) - ny
        if nw <= 0 or nh <= 0:
            return
        pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = cl, ct, cr, cb
        pic.left, pic.top, pic.width, pic.height = Emu(int(nx)), Emu(int(ny)), Emu(int(nw)), Emu(int(nh))

    def exact_math(self, slide, node, rec, under):
        svg = node.find("svg")
        srec = self.L(svg) or rec
        png = base64.b64decode(srec["png"].split(",", 1)[1]) if srec.get("png") else None   # rasterised by Chrome
        if png is None and svg is not None:
            png = render_svg_png(svg, srec["w"], srec["h"], css_to_hex(srec.get("c"), under) or "000000")
        if png:
            slide.shapes.add_picture(io.BytesIO(png), E(srec["x"]), E(srec["y"]), E(srec["w"]), E(srec["h"]))

    def exact_list(self, slide, lst, rec, under):
        items = []

        def walk(l, level):
            start = int(l["start"]) if str(l.get("start", "")).isdigit() else 1
            for li in l.find_all("li", recursive=False):
                lr = self.L(li)
                if lr is None:
                    continue
                runs, nested = [], []
                for c in li.children:
                    if isinstance(c, Tag) and c.name in ("ul", "ol"):
                        nested.append(c)
                        continue
                    if isinstance(c, Tag) and c.name in ("p", "div") and runs:
                        runs.append(Run(br=True))
                    self.exact_runs([c], lr, self.scale_of(lr), under, runs)
                runs = self.breaks(clean_runs(runs), lr)
                scale = self.scale_of(lr)
                x, top, _, own_h = self.content_box(lr)
                lh = self.line_px(lr, scale)
                bottom = top + own_h
                tops = [self.L(n)["y"] for n in nested if self.L(n)]
                if tops:
                    bottom = min(tops)
                top = self.text_top(lr, lh) if lr.get("ft") else top
                own_h = bottom - top
                if runs:
                    probe = {"runs": runs, "size": lr["fs"] * scale, "lh": lh}
                    items.append({"runs": runs, "x": x, "top": top, "own_h": own_h, "fs": lr["fs"] * scale,
                                  "baseline": css_baseline(lr, para_font(probe)[0]),
                                  "lh": self.line_px(lr, scale), "align": lr.get("ta", "left"),
                                  "style": lr.get("ls", "disc"), "color": css_to_hex(lr.get("mk"), under),
                                  "start": start})
                for n in nested:
                    walk(n, level + 1)

        walk(lst, 0)
        if not items:
            return
        box_x = rec["x"] + rec["b"][3]
        box_y = items[0]["top"]
        paras, prev_bottom = [], box_y
        for it in items:
            marL = it["x"] - box_x
            lines = max(1, round(it["own_h"] / it["lh"]))
            if self.forced_breaks:
                lines = 1 + sum(1 for r in it["runs"] if r.br)
            paras.append({"runs": it["runs"], "align": it["align"], "lh": it["lh"], "size": it["fs"],
                          "sb": it["top"] - prev_bottom, "lines": lines, "baseline": it["baseline"],
                          "bullet": {"marL": marL, "indent": -min(marL, it["fs"] * 1.0), "style": it["style"],
                                     "color": it["color"], "start": it["start"]}})
            prev_bottom = it["top"] + lines * it["lh"]
        self.exact_textbox(slide, paras, box_x, box_y, rec["x"] + rec["w"] - box_x, rec["y"] + rec["h"] - box_y,
                           name=lst.name)

    def exact_pre(self, slide, node, rec, under):
        self.decor(slide, node, rec, under)
        code = node.find("code")
        crec = self.L(code) or rec
        scale = self.scale_of(crec)
        runs = []
        self.exact_runs(node.children, crec, scale, under, runs, pre=True)
        while runs and (runs[-1].br or not runs[-1].text.strip()):
            runs.pop()
        if not runs:
            return
        x, y, w, h = self.content_box(rec)
        # the <pre> block's own line height (its strut) sets the pitch, not the inline <code>
        lh = max(self.line_px(rec, self.scale_of(rec)), self.line_px(crec, scale))
        para = {"runs": runs, "align": "left", "lh": lh, "size": crec["fs"] * scale}
        para["baseline"] = css_baseline(rec, para_font(para)[0])
        self.exact_textbox(slide, [para], x, y, w, h, name="Code")

    def exact_table(self, slide, table, rec, under):
        rows = [tr for tr in table.find_all("tr") if self.L(tr)]
        grid = [[c for c in tr.find_all(["th", "td"], recursive=False) if self.L(c)] for tr in rows]
        if not rows or not any(grid):
            return
        ncols = max(len(r) for r in grid)
        widest = max(grid, key=len)
        col_w = [self.L(c)["w"] for c in widest]
        total_w = sum(col_w)
        heights = [self.L(tr)["h"] for tr in rows]
        x, y = self.L(widest[0])["x"], self.L(rows[0])["y"]
        shape = slide.shapes.add_table(len(rows), ncols, E(x), E(y), E(total_w), E(sum(heights)))
        shape.name = "Table"
        tbl = shape.table
        tbl.first_row = tbl.horz_banding = False
        style = shape._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId"))
        if style is not None:
            style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"     # "No Style, No Grid"
        for ci, cw in enumerate(col_w):
            tbl.columns[ci].width = E(cw)
        for ri, (tr, cells) in enumerate(zip(rows, grid)):
            tbl.rows[ri].height = E(heights[ri])
            row_bg = css_to_hex(self.L(tr).get("bg"), under)
            for ci, cell_tag in enumerate(cells[:ncols]):
                cr = self.L(cell_tag)
                cell = tbl.cell(ri, ci)
                fill = css_to_hex(cr.get("bg"), under) or row_bg
                if fill:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb(fill)
                else:
                    cell.fill.background()
                pt_, pr_, pb_, pl_ = cr["p"]
                cell.margin_top, cell.margin_right, cell.margin_bottom, cell.margin_left = E(pt_), E(pr_), E(pb_), E(pl_)
                cell.vertical_anchor = {"top": MSO_ANCHOR.TOP, "bottom": MSO_ANCHOR.BOTTOM}.get(cr.get("va"), MSO_ANCHOR.MIDDLE)
                self.cell_borders(cell, cr, under)
                scale = self.scale_of(cr)
                runs = []
                self.exact_runs(cell_tag.children, cr, scale, under, runs)
                tf = cell.text_frame
                tf.word_wrap = True
                para = tf.paragraphs[0]
                para.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(cr.get("ta"), PP_ALIGN.LEFT)
                para.line_spacing = ppt_pct(cr["fs"] * scale, self.line_px(cr, scale))
                for r in self.breaks(clean_runs(runs), cr):
                    if r.br:
                        para.add_line_break()
                        continue
                    run = para.add_run()
                    run.text = r.text
                    run.font.size = PT(r.size)
                    run.font.name = r.font
                    run.font.bold = True if r.bold else None
                    run.font.italic = True if r.italic else None
                    if r.color:
                        run.font.color.rgb = rgb(r.color)

    @staticmethod
    def cell_borders(cell, cr, under):
        tcPr = cell._tc.get_or_add_tcPr()
        for el in list(tcPr):
            if el.tag in (qn("a:lnL"), qn("a:lnR"), qn("a:lnT"), qn("a:lnB")):
                tcPr.remove(el)
        sides = {"lnL": 3, "lnR": 1, "lnT": 0, "lnB": 2}
        for pos, tag in enumerate(("lnL", "lnR", "lnT", "lnB")):     # schema order, before the fill
            i = sides[tag]
            width = cr["b"][i]
            colour = css_to_hex(cr["bc"][i], under) if cr.get("bc") else None
            ln = etree.Element(qn(f"a:{tag}"))
            if width >= 0.5 and colour:
                ln.set("w", str(int(E(width))))
                etree.SubElement(etree.SubElement(ln, qn("a:solidFill")), qn("a:srgbClr")).set("val", colour)
            else:
                ln.set("w", "0")
                etree.SubElement(ln, qn("a:noFill"))
            tcPr.insert(pos, ln)

    def emit(self, slide, node, under):
        """Place one element (and its descendants) at the coordinates Chrome measured."""
        if not isinstance(node, Tag):
            return
        rec = self.L(node)
        name = node.name
        if rec is None or name in ("style", "script", "br"):
            return
        if name == "img":
            return self.exact_image(slide, node, rec, under)
        if name == "video":
            return self.video(slide, {"src": node.get("src", "")}, rec["x"], rec["y"], rec["w"], rec["h"],
                              {"card_bg": "F4F4F4", "bg": under})
        if name == "mjx-container":
            return self.exact_math(slide, node, rec, under)
        if name == "svg":
            if node.has_attr("data-marp-fitting"):        # Marp auto-fit text, scaled by the SVG
                return self.exact_text(slide, node, rec, under)
            return
        if name == "table":
            return self.exact_table(slide, node, rec, under)
        if name == "pre":
            return self.exact_pre(slide, node, rec, under)
        if name in ("ul", "ol"):
            self.decor(slide, node, rec, under)
            self.exact_list(slide, node, rec, under)
            for media in node.find_all(["img", "video", "mjx-container"]):
                self.emit(slide, media, under)
            return
        self.decor(slide, node, rec, under)
        # <br> is blockified inside flex/grid containers but still just breaks the line
        blocks = [c for c in node.children if isinstance(c, Tag) and c.name != "br" and self.L(c)
                  and self.L(c).get("d") not in INLINE_DISPLAYS]
        if blocks:
            loose = [c for c in node.children if c not in blocks]
            if any(isinstance(c, NavigableString) and not isinstance(c, Comment) and c.strip() for c in loose):
                self.exact_text(slide, node, rec, under, nodes=loose)     # text beside block children
            for c in node.children:
                self.emit(slide, c, under)
        else:
            self.exact_text(slide, node, rec, under)
            for media in node.find_all(["img", "video", "mjx-container"]):
                self.emit(slide, media, under)

    def build_exact(self, slide, section, pal):
        under = css_to_hex(self.layout.get("slide" + section["data-m2p-slide"], {}).get("bg")) or pal["bg"]
        for child in section.children:
            self.emit(slide, child, under)

    def footer(self, slide, text, pal, draw_rule=True):
        light = luminance(pal["bg"]) > 0.5
        if draw_rule:   # the rendered theme background already contains the rule
            line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(PAD_X), E(CANVAS_H - 60), E(CANVAS_W - PAD_X), E(CANVAS_H - 60))
            line.line.color.rgb = rgb(RULE_COLOR if light else pal["border"])
            line.line.width = Pt(0.75)
            line.name = "Footer rule"
        if text:
            color = pal.get("footer") or (SU["cool_grey"] if light else "D9D9D9")
            p = Para([Run(text)], size=16, color=color, align="center", space_after=0, line=1.0)
            self.textbox(slide, [p], PAD_X, CANVAS_H - 46, CONTENT_W, 24, 1, pal, anchor="middle", name="Footer")


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------
def parse_range(spec: str | None, n: int) -> list:
    if not spec:
        return list(range(1, n + 1))
    keep = set()
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            keep.update(range(int(a), int(b) + 1))
        elif part.strip():
            keep.add(int(part))
    return [i for i in range(1, n + 1) if i in keep]


def environment(chrome_path: str | None = None) -> dict:
    """What the converter can use on this machine (shown by the app)."""
    chrome = find_chrome(chrome_path)
    fonts = office_font_css()
    return {"chrome": chrome, "calibri": "Calibri" in fonts, "cairosvg": cairosvg is not None}


def convert(html: str, output: str | None = None, *, slides: str | None = None, download: bool = True,
            video: bool = True, max_image_px: int = 2560, backgrounds: bool = True, chrome: str | None = None,
            natural_wrap: bool = False, approx_layout: bool = False, progress=None, log=None) -> dict:
    """Convert a Marp HTML deck to .pptx and return a summary.

    progress(stage, done, total) is called with stage in "read", "backgrounds", "layout", "slides",
    "save", "done"; log(message) receives warnings and status lines."""
    global _LOG
    _LOG = log
    step = progress or (lambda *a: None)
    try:
        html_path = os.path.abspath(html)
        out = output or os.path.splitext(html_path)[0] + ".pptx"
        step("read", 0, 1)
        with open(html_path, encoding="utf-8") as fh:
            soup = BeautifulSoup(fh.read(), "lxml")
        sections = [s for s in soup.find_all("section") if s.find_parent("svg") is not None and s.get("id")
                    and not s.has_attr("data-marpit-advanced-background")]
        if not sections:
            raise ValueError("No Marp slides found in this HTML file (expected a Marp export).")
        notes = {}
        for div in soup.select("div.bespoke-marp-note"):
            idx = int(div.get("data-index", -1))
            notes[idx] = "\n".join(p.get_text("\n").strip() for p in div.find_all(["p", "li"])) or div.get_text("\n").strip()

        wanted = parse_range(slides, len(sections))
        themes, layout = {}, {}
        chrome_bin = find_chrome(chrome)
        css = "\n".join(st.get_text() for st in soup.find_all("style"))   # the deck's compiled theme CSS
        selected = [sections[i - 1] for i in wanted]
        if chrome_bin:
            tag_sections(selected)
            if backgrounds:
                step("backgrounds", 0, 1)
                themes = ThemeRenderer(chrome_bin, css).render(selected)
            if not approx_layout:
                step("layout", 0, 1)
                layout = measure_layout(chrome_bin, css, selected, os.path.dirname(html_path))
        else:
            warn("Chrome not found: flat backgrounds and approximate layout (install Google Chrome to fix)")

        images = ImageStore(os.path.dirname(html_path), download=download, max_px=max_image_px)
        deck = DeckBuilder(images, embed_video=video, themes=themes, layout=layout)
        deck.forced_breaks = not natural_wrap
        warn(f"{len(sections)} slides found, converting {len(wanted)} ...")
        failed = []
        for n, i in enumerate(wanted, 1):
            try:
                deck.build(sections[i - 1], notes.get(i - 1, ""))
            except Exception as exc:
                failed.append(i)
                warn(f"  ! slide {i} failed ({exc}); left blank")
            step("slides", n, len(wanted))
        step("save", 0, 1)
        deck.prs.save(out)
        warn(f"Saved {out}")
        missing = list(dict.fromkeys(images.missing))
        if missing:
            warn(f"{len(missing)} image(s) could not be loaded (shown as placeholders):")
            for m in missing:
                warn("  - " + m)
        summary = {"output": out, "slides": len(wanted), "failed": failed, "missing_images": missing,
                   "exact_layout": bool(layout), "backgrounds": bool(themes), "size": os.path.getsize(out)}
        step("done", 1, 1)
        return summary
    finally:
        _LOG = None


def main():
    ap = argparse.ArgumentParser(description="Convert a Marp HTML export into an editable PowerPoint deck.")
    ap.add_argument("html", help="Marp-exported .html file")
    ap.add_argument("-o", "--output", help="output .pptx (default: next to the HTML)")
    ap.add_argument("--slides", help="subset to convert, e.g. 1-40,120")
    ap.add_argument("--no-download", action="store_true", help="do not fetch http(s) images")
    ap.add_argument("--no-video", action="store_true", help="placeholder instead of embedding videos")
    ap.add_argument("--max-image-px", type=int, default=2560, help="downscale larger images (0 = keep originals)")
    ap.add_argument("--no-backgrounds", action="store_true", help="flat colours instead of Chrome-rendered theme backgrounds")
    ap.add_argument("--chrome", help="path to Chrome/Chromium/Edge (auto-detected by default)")
    ap.add_argument("--natural-wrap", action="store_true",
                    help="let PowerPoint wrap lines itself instead of copying the browser's line breaks")
    ap.add_argument("--approx-layout", action="store_true",
                    help="use the built-in approximate layout instead of Chrome-measured positions")
    args = ap.parse_args()

    def progress(stage, done, total):
        if stage == "slides" and (done % 100 == 0 or done == total):
            print(f"  {done}/{total}", file=sys.stderr)

    try:
        convert(args.html, args.output, slides=args.slides, download=not args.no_download, video=not args.no_video,
                max_image_px=args.max_image_px, backgrounds=not args.no_backgrounds, chrome=args.chrome,
                natural_wrap=args.natural_wrap, approx_layout=args.approx_layout, progress=progress)
    except (OSError, ValueError) as exc:
        sys.exit(f"Error: {exc}")


if __name__ == "__main__":
    main()
