#!/usr/bin/env bash
# Start CareerTrace Pro (no key prompts - manage keys inside the app: Settings > API keys)
cd "$(dirname "$0")"
[ -f .env ] && set -a && . ./.env && set +a
[ -f .env.pro ] && set -a && . ./.env.pro && set +a
PY=$(command -v python3 || command -v python)
echo "Starting CareerTrace Pro. In Codespaces open the Ports tab > port 8000 > globe icon if it does not open by itself."
exec $PY desktop.py
