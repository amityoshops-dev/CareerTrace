#!/usr/bin/env bash
# CareerTrace Pro v2 installer - run from the repo root:  bash CareerTrace-Pro-v2/install.sh
set -e
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
[ -f careertrace_ui.py ] || { echo "Run this from the CareerTrace repo root (where careertrace_ui.py is)."; exit 1; }
PY=$(command -v python3 || command -v python)
[ -f careertrace_ui.py.bak_pro ] || cp careertrace_ui.py careertrace_ui.py.bak_pro
mkdir -p data exports pro_static output
$PY -m pip install -q openpyxl python-docx pywebview 2>/dev/null || $PY -m pip install -q --break-system-packages openpyxl python-docx 2>&1 | tail -1 || true
cp "$HERE/careertrace_pro.py" ./careertrace_pro.py
cp "$HERE/pro_static/"* ./pro_static/
cp "$HERE/desktop.py" "$HERE/start_pro.sh" "$HERE/start_windows.bat" ./
chmod +x start_pro.sh
for f in .env .env.pro data/ exports/ output/ '*.bak_pro'; do
  grep -qxF "$f" .gitignore 2>/dev/null || echo "$f" >> .gitignore
done
if ! grep -q "careertrace_pro" careertrace_ui.py; then
cat >> careertrace_ui.py <<'LOADEOF'


# --- CareerTrace Pro (safe loader: a failure here never stops the app) ---
try:
    import careertrace_pro
    careertrace_pro.install(app, globals())
except Exception as _e:
    print("CareerTrace Pro not loaded:", _e)
LOADEOF
fi
$PY -m py_compile careertrace_pro.py
if $PY -c "import careertrace_ui as u; assert '/api/pro/health' in u.app.openapi()['paths']" >/tmp/pro_check.log 2>&1; then
  echo "OK: CareerTrace Pro v2 installed. Start it with:  ./start_pro.sh"
else
  cp careertrace_ui.py.bak_pro careertrace_ui.py
  echo "Install check failed, your app was restored unchanged. Details:"; tail -15 /tmp/pro_check.log
  exit 1
fi
