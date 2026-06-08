#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."
BRANCH="${RTWEB_UPDATE_BRANCH:-main}"
git fetch origin "$BRANCH"
git pull --ff-only origin "$BRANCH"
python3 -m py_compile app.py
TMPDB="$(mktemp /tmp/ms-mail-manager-test-db.XXXXXX)"
rm -f "$TMPDB"
RTWEB_DB="$TMPDB" python3 -m unittest discover -s tests -p 'test_*.py'
rm -f "$TMPDB"
if command -v systemctl >/dev/null 2>&1; then
  systemctl restart "${RTWEB_SERVICE:-rtweb}"
fi
