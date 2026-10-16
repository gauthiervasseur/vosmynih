marp_html_to_pptx.py - convert a Marp HTML export into an editable PowerPoint
==============================================================================

1. ONE-TIME SETUP (from this folder)
------------------------------------
cd ~/Shared-SynologyDrive/SynologyDrive/Code/marp2pptx
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

(cairosvg is optional: it renders MathJax equations. If it fails to install,
equations show as "[equation]" and everything else still works.)


2. EACH TIME YOU CONVERT
------------------------
cd ~/Shared-SynologyDrive/SynologyDrive/Code/marp2pptx
source .venv/bin/activate

# Basic: writes deck.pptx next to deck.html
python marp_html_to_pptx.py "/path/to/deck.html"

# Example with the TECH78 deck
python marp_html_to_pptx.py "/Users/gauthiervasseur/Shared-SynologyDrive/Markdown Dossiers/DWA1/Engagements/DWA-Base/2026-10-24-TECH78-Data_BI-AI-GenAI.html"

# Choose the output file
python marp_html_to_pptx.py deck.html -o ~/Desktop/deck.pptx

# Convert only some slides (handy for testing)
python marp_html_to_pptx.py deck.html --slides 1-40,120,300-310

# Smaller file: don't embed videos (puts a link to the video file instead)
python marp_html_to_pptx.py deck.html --no-video

# Offline: skip downloading http(s) images (shown as placeholders)
python marp_html_to_pptx.py deck.html --no-download

# Keep images at full resolution (default downscales anything over 2560 px)
python marp_html_to_pptx.py deck.html --max-image-px 0

# Google Chrome (found automatically) lays out every slide with the deck's CSS, so
# PowerPoint gets the exact positions, sizes, colours and line breaks of the HTML,
# plus the theme backgrounds (colour, mosaic, Stanford wordmark, footer rule).
# If Chrome lives elsewhere:
python marp_html_to_pptx.py deck.html --chrome "/Applications/Chromium.app/Contents/MacOS/Chromium"

# Let PowerPoint wrap lines itself (no soft line breaks inserted; easier to rewrite
# paragraphs later, but a line may occasionally wrap differently from the HTML)
python marp_html_to_pptx.py deck.html --natural-wrap

# Flat background colours instead of the rendered theme backgrounds
python marp_html_to_pptx.py deck.html --no-backgrounds

# Without Chrome: approximate built-in layout (positions are estimated, not exact)
python marp_html_to_pptx.py deck.html --approx-layout --no-backgrounds

# Show all options
python marp_html_to_pptx.py --help


3. NOTES
--------
- Always quote paths that contain spaces ("Markdown Dossiers/...").
- Relative image paths (../../res/...) are resolved from the HTML file's folder,
  so run it on the HTML in its original location.
- An existing .pptx with the same name is overwritten.
- Images that can't be found appear as grey "Image not found" boxes and are
  listed at the end of the run.
- Alignment: Chrome measures text with the Calibri that ships with Microsoft
  Office (inside PowerPoint.app), i.e. the font PowerPoint will draw. Text lands
  within about 1 px of the HTML. Keep Office installed on the machine that runs it.
- Wrapped lines contain soft line breaks (Shift+Enter) copied from the browser.
  If you rewrite a paragraph, delete them, or convert with --natural-wrap.
- Code blocks are drawn in Consolas (as the CSS asks); your browser shows them in
  Menlo because Consolas isn't installed system-wide, so code looks a bit narrower.


4. THE ВОСЬМИНІГ APP (retro arcade desktop app)
------------------------------------------------
# Run it from source (dev)
pip install -r requirements-app.txt
python app/vosmynih_app.py

# Build the Mac app + DMG  ->  dist/vosmynih-<version>-macos-arm64.dmg
#   (build environment is kept in ~/.cache/vosmynih, outside this synced folder)
./build/build_macos.sh /opt/homebrew/bin/python3.13

# Same, plus an end-to-end self-test of the built app on a real deck
SELFTEST_DECK="/path/to/deck.html" ./build/build_macos.sh /opt/homebrew/bin/python3.13

# Build for Windows (run ON a Windows PC; makes a zip, and a setup.exe if Inno Setup 6 is installed)
powershell -ExecutionPolicy Bypass -File build\build_windows.ps1

# Build both platforms in the cloud: copy build/github-actions-build.yml to
# <repo>/.github/workflows/ and push a tag like vosmynih-v1.0.0

# Regenerate the octopus sprite, sound and icons
python tools/make_assets.py


5. THE WEBSITE
--------------
# Assemble site/ from web/index.html + the installers found in dist/
python tools/build_site.py

# Preview locally (the site is plain static files)
python -m http.server 8000 --directory site      # then open http://localhost:8000

# Publish: upload the whole site/ folder to any static host
#   (Synology Web Station, GitHub Pages, Netlify...). index.html?silent skips the intro screen.
