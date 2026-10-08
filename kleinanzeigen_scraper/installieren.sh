#!/usr/bin/env bash
# Installiert alles, was der Scraper braucht (macOS / Linux).
set -e
cd "$(dirname "$0")"
echo "=== Kleinanzeigen-Scraper: Installation ==="

if ! command -v python3 >/dev/null 2>&1; then
  if [[ "$OSTYPE" == darwin* ]] && command -v brew >/dev/null 2>&1; then
    echo "Python 3 fehlt – Installation über Homebrew ..."
    brew install python
  else
    echo "Python 3 fehlt. Bitte von https://www.python.org/downloads/ installieren und erneut starten."
    exit 1
  fi
fi

[ -x .venv/bin/python ] || python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip --quiet
.venv/bin/python -m pip install -r requirements.txt
echo
echo "Installation fertig. Starten mit ./starten.sh --test oder ./starten.sh"
