#!/bin/zsh
# Build восьминіг.app and a drag-to-install DMG.
#   ./build/build_macos.sh [path/to/python]         (needs: python 3.10+, internet for pip the first time)
# Output: dist/vosmynih-<version>-macos-<arch>.dmg
set -euo pipefail
ROOT=${0:A:h:h}
cd "$ROOT"
PY=${1:-python3}
VERSION=$(sed -n 's/^VERSION = "\(.*\)"/\1/p' app/vosmynih_app.py)
ARCH=$(uname -m)
APP="восьминіг"

CACHE=${VOSMYNIH_CACHE:-$HOME/.cache/vosmynih}       # kept out of the (synced) source folder
echo "==> build environment in $CACHE"
mkdir -p "$CACHE"
[[ -d "$CACHE/build-venv" ]] || "$PY" -m venv "$CACHE/build-venv"
source "$CACHE/build-venv/bin/activate"
pip -q install --upgrade pip
pip -q install -r requirements-app.txt pyinstaller fonttools

echo "==> assets"
python tools/make_assets.py

echo "==> bundle"
rm -rf "$CACHE/pyinstaller" dist/vosmynih.app "dist/$APP.app"
pyinstaller --noconfirm --clean --windowed --name vosmynih \
  --icon "$ROOT/app/assets/icon.icns" --osx-bundle-identifier com.datawiseacademy.vosmynih \
  --add-data "$ROOT/app/ui:ui" --paths "$ROOT" --collect-data pptx \
  --exclude-module cairosvg --exclude-module cairocffi --exclude-module tkinter \
  --workpath "$CACHE/pyinstaller" --specpath "$CACHE/pyinstaller" --distpath dist \
  "$ROOT/app/vosmynih_app.py" >/dev/null

mv dist/vosmynih.app "dist/$APP.app"
rm -rf dist/vosmynih                                 # PyInstaller's intermediate folder
python - "$ROOT/dist/$APP.app/Contents/Info.plist" "$APP" "$VERSION" <<'PY'
import plistlib, sys
path, name, version = sys.argv[1:]
with open(path, "rb") as fh:
    info = plistlib.load(fh)
info.update(CFBundleName=name, CFBundleDisplayName=name, CFBundleShortVersionString=version, CFBundleVersion=version,
            NSHumanReadableCopyright="восьминіг - Marp HTML to PowerPoint", NSHighResolutionCapable=True,
            LSMinimumSystemVersion="11.0")
with open(path, "wb") as fh:
    plistlib.dump(info, fh)
PY
codesign --force --deep --sign - "dist/$APP.app"      # ad-hoc signature (required on Apple Silicon)

if [[ -n "${SELFTEST_DECK:-}" ]]; then
  echo "==> self-test with $SELFTEST_DECK"
  "dist/$APP.app/Contents/MacOS/vosmynih" --selftest "$SELFTEST_DECK" "$ROOT/dist/selftest.pptx" | grep SELFTEST
  rm -f dist/selftest.pptx
fi

echo "==> dmg"
STAGE=$(mktemp -d)
cp -R "dist/$APP.app" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
cat > "$STAGE/READ ME FIRST.txt" <<TXT
восьминіг $VERSION - Marp HTML deck to editable PowerPoint

1. Drag восьминіг onto Applications.
2. First launch: right-click the app > Open > Open (the app is not notarised by Apple).
3. Needs Google Chrome (exact layout and theme backgrounds) and Microsoft Office
   (its Calibri font makes text land where the HTML puts it).
TXT
DMG="dist/vosmynih-$VERSION-macos-$ARCH.dmg"
rm -f "$DMG"
hdiutil create -volname "$APP" -srcfolder "$STAGE" -ov -format UDZO "$DMG" >/dev/null 2>&1
rm -rf "$STAGE"
echo "==> done: $DMG ($(du -h "$DMG" | cut -f1))"
