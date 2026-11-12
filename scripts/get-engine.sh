#!/usr/bin/env bash
# Builds the editor engine from its open-source project and applies the YourOfficeSuite changes.
# Needs: git, Node.js 24 or newer, corepack (ships with Node), python3. Takes a few minutes and about 2 GB of disk.
set -euo pipefail
ENGINE_REPO="https://github.com/ranuts/document.git"
ENGINE_COMMIT="d1c9b03263aabe14789297f5c717fefff2fe4f0d"   # the commit YourOfficeSuite 2.2.0 was built and tested with
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="$ROOT/.engine-src"

node -e 'if (+process.versions.node.split(".")[0] < 24) { console.error("Node.js 24 or newer is needed to build the engine"); process.exit(1) }'
if [ ! -d "$WORK/.git" ]; then git clone "$ENGINE_REPO" "$WORK"; fi
cd "$WORK"
git fetch --depth 1 origin "$ENGINE_COMMIT" 2>/dev/null || git fetch origin
git checkout -q "$ENGINE_COMMIT"
corepack pnpm@11.4.0 install --frozen-lockfile
mkdir -p .bin && printf '#!/bin/sh\nexec corepack pnpm@11.4.0 "$@"\n' > .bin/pnpm && chmod +x .bin/pnpm
PATH="$WORK/.bin:$PATH" sh ./bin/build.sh
rm -rf "$ROOT/engine" && cp -r dist "$ROOT/engine"
python3 "$ROOT/scripts/patch-engine.py" "$ROOT/engine"
echo "Engine ready in $ROOT/engine"
