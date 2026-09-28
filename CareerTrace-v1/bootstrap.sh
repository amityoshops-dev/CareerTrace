#!/usr/bin/env bash
set -euo pipefail
echo "== CareerTrace v1 bootstrap =="
python3 --version
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
mkdir -p data
echo "Bootstrap complete."
echo "Set GOOGLE_API_KEY and LANGSMITH_API_KEY as Codespaces secrets."
echo "Put resume at data/resume.docx."
echo "Run: source .venv/bin/activate && uvicorn app.main:app --host 0.0.0.0 --port 8000"
