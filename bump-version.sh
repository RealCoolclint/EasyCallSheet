#!/usr/bin/env bash
# Met à jour la date de version (DD.MM.YY) avant un commit / push de correctif.
# Usage : ./bump-version.sh

set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
DATE=$(date +%d.%m.%y)
VERSION="V1.30.${DATE}"

sed -i '' "s/window\.APP_VERSION = 'V1\.30\.[^']*'/window.APP_VERSION = '${VERSION}'/" "${ROOT}/version.js"
sed -i '' "s/V1\.30\.[0-9.]* ·/V1.30.${DATE} ·/" "${ROOT}/styles.css"

echo "Version → ${VERSION} (version.js + styles.css)"
