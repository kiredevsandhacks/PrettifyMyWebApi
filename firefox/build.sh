#!/usr/bin/env bash
# Builds the Firefox version of the extension into dist/firefox and packs it into dist/*.zip
# Usage: ./firefox/build.sh
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
out="$root/dist/firefox"

rm -rf "$out"
mkdir -p "$out"

# shared source, without the Chrome manifest and without the packaged Chrome zip
(cd "$root/src" && tar --exclude='./manifest.json' --exclude='./prettifyWebApi.zip' -cf - .) | (cd "$out" && tar -xf -)
cp "$root/firefox/manifest.json" "$out/manifest.json"

# eventHandler.js is injected by background.js, it does not need to be web accessible
npx --yes web-ext lint --source-dir "$out" || true
npx --yes web-ext build --source-dir "$out" --artifacts-dir "$root/dist" --overwrite-dest
