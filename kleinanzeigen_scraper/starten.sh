#!/usr/bin/env bash
# Startet den Scraper (Standard: 500 Häuser). Optionen werden durchgereicht,
# z.B.:  ./starten.sh --test   oder   ./starten.sh --ziel 200 --nur-kauf
set -e
cd "$(dirname "$0")"
[ -x .venv/bin/python ] || ./installieren.sh
exec .venv/bin/python kleinanzeigen_haeuser.py "$@"
