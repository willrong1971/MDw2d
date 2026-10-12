#!/bin/bash
set -euo pipefail

usage() {
  cat <<'USAGE'
Build the self-contained Apple Silicon macOS app.
Usage: build.sh --destination DIRECTORY [--converter PATH] [--python-home PATH] [--libreoffice-app PATH] [--icon PATH] [--full-python]
       build.sh --compile-only
The destination must not already contain MDw2d.app. Existing apps are never overwritten.
USAGE
}

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONVERTER="${CONVERTER_SOURCE:-$SCRIPT_DIR/../engine/converter.py}"
DESTINATION=""
PYTHON_SOURCE="${PYTHON_HOME:-$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python}"
LIBREOFFICE_SOURCE="${LIBREOFFICE_APP:-$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/native/libreoffice-headless/libreoffice/LibreOfficeDev.app}"
ICON_SOURCE="${APP_ICON:-$SCRIPT_DIR/../assets/AppIcon.icns}"
FULL_PYTHON=0
COMPILE_ONLY=0
while [[ $# -gt 0 ]]; do
  case "$1" in
    --converter) CONVERTER="$2"; shift 2 ;;
    --destination) DESTINATION="$2"; shift 2 ;;
    --python-home) PYTHON_SOURCE="$2"; shift 2 ;;
    --libreoffice-app) LIBREOFFICE_SOURCE="$2"; shift 2 ;;
    --icon) ICON_SOURCE="$2"; shift 2 ;;
    --full-python) FULL_PYTHON=1; shift ;;
    --compile-only) COMPILE_ONLY=1; shift ;;
    --help|-h) usage; exit 0 ;;
    *) usage >&2; exit 2 ;;
  esac
done

mkdir -p "$SCRIPT_DIR/build"
SWIFT_BINARY="$SCRIPT_DIR/build/MDw2d"
xcrun swiftc -O -parse-as-library -swift-version 5 -target arm64-apple-macos14.0 \
  -framework AppKit -framework SwiftUI -framework UniformTypeIdentifiers \
  -o "$SWIFT_BINARY" "$SCRIPT_DIR/Sources/ConverterApp.swift"
if [[ "$COMPILE_ONLY" == 1 ]]; then
  printf 'Compiled: %s\n' "$SWIFT_BINARY"
  exit 0
fi
if [[ -z "$CONVERTER" || -z "$DESTINATION" ]]; then usage >&2; exit 2; fi
[[ -f "$CONVERTER" ]] || { printf 'Converter missing: %s\n' "$CONVERTER" >&2; exit 1; }
[[ -x "$PYTHON_SOURCE/bin/python3" ]] || { printf 'Python missing: %s\n' "$PYTHON_SOURCE" >&2; exit 1; }
[[ -x "$LIBREOFFICE_SOURCE/Contents/MacOS/soffice" ]] || { printf 'LibreOffice missing: %s\n' "$LIBREOFFICE_SOURCE" >&2; exit 1; }

mkdir -p "$DESTINATION"
DESTINATION="$(cd "$DESTINATION" && pwd)"
APP="$DESTINATION/MDw2d.app"
[[ ! -e "$APP" ]] || { printf 'Refusing to overwrite existing app: %s\n' "$APP" >&2; exit 1; }
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources/runtime" "$APP/Contents/Resources/libreoffice"
cp "$SWIFT_BINARY" "$APP/Contents/MacOS/MDw2d"
cp "$SCRIPT_DIR/Info.plist" "$APP/Contents/Info.plist"
if [[ -f "$ICON_SOURCE" ]]; then
  cp "$ICON_SOURCE" "$APP/Contents/Resources/AppIcon.icns"
  /usr/libexec/PlistBuddy -c 'Add :CFBundleIconFile string AppIcon' "$APP/Contents/Info.plist"
fi
cp "$CONVERTER" "$APP/Contents/Resources/converter.py"
CONVERTER_DIR="$(cd "$(dirname "$CONVERTER")" && pwd)"
if [[ -f "$CONVERTER_DIR/content_audit.py" ]]; then
  cp "$CONVERTER_DIR/content_audit.py" "$APP/Contents/Resources/content_audit.py"
fi
cp "$SCRIPT_DIR/THIRD-PARTY-NOTICES.md" "$APP/Contents/Resources/THIRD-PARTY-NOTICES.md"

if [[ "$FULL_PYTHON" == 1 ]]; then
  ditto "$PYTHON_SOURCE" "$APP/Contents/Resources/runtime/python"
else
  # Keep CPython and its standard library but omit the unrelated artifact
  # processing packages in the shared runtime. Preserve installed licenses.
  mkdir -p "$APP/Contents/Resources/runtime/python/bin" "$APP/Contents/Resources/runtime/python/lib"
  PYTHON_EXECUTABLE="$PYTHON_SOURCE/bin/python3.12"
  [[ -x "$PYTHON_EXECUTABLE" ]] || { printf 'Minimal bundle expects CPython 3.12.\n' >&2; exit 1; }
  cp "$PYTHON_EXECUTABLE" "$APP/Contents/Resources/runtime/python/bin/python3.12"
  ln -s python3.12 "$APP/Contents/Resources/runtime/python/bin/python3"
  rsync -a --exclude site-packages --exclude __pycache__ --exclude test \
    "$PYTHON_SOURCE/lib/python3.12/" "$APP/Contents/Resources/runtime/python/lib/python3.12/"
  if [[ -d "$PYTHON_SOURCE/share" ]]; then
    ditto "$PYTHON_SOURCE/share" "$APP/Contents/Resources/runtime/python/share"
  fi
  if [[ -f "$PYTHON_SOURCE/lib/libpython3.12.dylib" ]]; then
    cp "$PYTHON_SOURCE/lib/libpython3.12.dylib" "$APP/Contents/Resources/runtime/python/lib/"
  fi
fi

# The supplied runtime's man-page alias points at its former build directory.
# Keep it local so the delivered bundle has no external or dangling links.
PYTHON_MAN_DIR="$APP/Contents/Resources/runtime/python/share/man/man1"
if [[ -L "$PYTHON_MAN_DIR/python3.1" && -f "$PYTHON_MAN_DIR/python3.12.1" ]]; then
  ln -sfn python3.12.1 "$PYTHON_MAN_DIR/python3.1"
fi

# Keep the whole upstream application, including notices, filter data, fonts,
# and export dependencies. The prototype never uses a user's LibreOffice.
ditto "$LIBREOFFICE_SOURCE" "$APP/Contents/Resources/libreoffice/LibreOfficeDev.app"
plutil -lint "$APP/Contents/Info.plist"

# The local prototype is ad-hoc signed. The embedded runtimes were patched by
# their provider; re-signing the copied app checks the actual delivered files.
codesign --force --deep --sign - "$APP"
codesign --verify --deep --strict "$APP"
"$APP/Contents/Resources/runtime/python/bin/python3" -I -c 'import json, pathlib, hashlib, struct, subprocess, tempfile, zipfile, xml.etree.ElementTree; print("Bundled Python: OK")'
printf 'Built: %s\n' "$APP"
du -sh "$APP"
