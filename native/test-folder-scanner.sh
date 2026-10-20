#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TEST_DIRECTORY="$(mktemp -d "${TMPDIR:-/tmp}/MDw2d-folder-tests.XXXXXX")"
trap 'rm -rf "$TEST_DIRECTORY"' EXIT

xcrun swiftc -O -parse-as-library -swift-version 5 -target arm64-apple-macos14.0 \
  "$SCRIPT_DIR/Sources/FolderScanner.swift" "$SCRIPT_DIR/Tests/FolderScannerTests.swift" \
  -o "$TEST_DIRECTORY/FolderScannerTests"
"$TEST_DIRECTORY/FolderScannerTests"
