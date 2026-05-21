#!/usr/bin/env bash
# Met à jour la version (DD.MM.YY) et optionnellement publie la release GitHub.
#
# Usage :
#   ./bump-version.sh              → met à jour version.js + styles.css
#   ./bump-version.sh --release    → tag + release GitHub (après commit & push)

set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"

read_version() {
    grep "window.APP_VERSION" "${ROOT}/version.js" | sed "s/.*'\([^']*\)'.*/\1/"
}

bump_files() {
    local DATE VERSION
    DATE=$(date +%d.%m.%y)
    VERSION="V1.30.${DATE}"

    sed -i '' "s/window\.APP_VERSION = 'V1\.30\.[^']*'/window.APP_VERSION = '${VERSION}'/" "${ROOT}/version.js"
    sed -i '' "s/V1\.30\.[0-9.]* ·/V1.30.${DATE} ·/" "${ROOT}/styles.css"

    echo "Version → ${VERSION} (version.js + styles.css)"
    echo "Ensuite : git add … && git commit && git push origin main"
    echo "Puis    : ./bump-version.sh --release"
}

publish_release() {
    local VERSION NOTES
    VERSION=$(read_version)

    if git rev-parse "${VERSION}" >/dev/null 2>&1; then
        echo "Le tag ${VERSION} existe déjà."
    else
        git tag -a "${VERSION}" -m "Manifest ${VERSION}"
        git push origin "${VERSION}"
    fi

    if gh release view "${VERSION}" >/dev/null 2>&1; then
        echo "La release ${VERSION} existe déjà sur GitHub."
        return 0
    fi

    NOTES="${2:-}"
    if [[ -z "${NOTES}" ]]; then
        NOTES="Release ${VERSION}."
    fi

    gh release create "${VERSION}" \
        --title "Manifest ${VERSION}" \
        --notes "${NOTES}"

    echo "Release publiée : https://github.com/RealCoolclint/Manifest/releases/tag/${VERSION}"
}

case "${1:-}" in
    --release|-r)
        publish_release
        ;;
    "")
        bump_files
        ;;
    *)
        echo "Usage: $0 | $0 --release"
        exit 1
        ;;
esac
