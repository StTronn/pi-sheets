#!/usr/bin/env bash
# install.sh — one-shot setup for pi-sheets python deps.
#
# Flow:
#   1. Fast-exit if the cached venv already has the deps.
#   2. Sanity-check `python3` (must be >= 3.10).
#   3. Create venv at $XDG_DATA_HOME/pi-sheets/.venv (~/.local/share/pi-sheets/.venv).
#   4. Install openpyxl + formualizer.
#   5. Cache the venv-python path so xlsx.py auto-redirects there.
#
# Idempotent. Errors surface plainly; this script doesn't try to be clever about
# broken pythons — if `python3 -m venv` fails, fix python and re-run.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REQS="$HERE/../requirements.txt"

VENV="${XDG_DATA_HOME:-$HOME/.local/share}/pi-sheets/.venv"
STATE="${XDG_STATE_HOME:-$HOME/.local/state}/pi-sheets/python.txt"

ok()  { printf "\033[32m[pi-sheets]\033[0m %s\n" "$*"; }
err() { printf "\033[31m[pi-sheets]\033[0m %s\n" "$*" >&2; }

# 1. Fast-exit if cached interpreter still has both deps.
if [ -f "$STATE" ]; then
  CACHED="$(cat "$STATE")"
  if [ -x "$CACHED" ] && "$CACHED" -c "import openpyxl, formualizer" 2>/dev/null; then
    ok "ready — $CACHED"
    exit 0
  fi
fi

# 2. Need python3 >= 3.10.
if ! command -v python3 >/dev/null 2>&1; then
  err "python3 not found on PATH."
  err "Install: curl -LsSf https://astral.sh/uv/install.sh | sh && uv python install"
  exit 1
fi
PY_VER=$(python3 -c "import sys; print('%d.%d' % sys.version_info[:2])")
if ! python3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"; then
  err "python3 is $PY_VER; pi-sheets needs >= 3.10."
  err "Install: curl -LsSf https://astral.sh/uv/install.sh | sh && uv python install"
  exit 1
fi

# 3. Create the venv (clobbers any stale one).
mkdir -p "$(dirname "$VENV")" "$(dirname "$STATE")"
rm -rf "$VENV"
python3 -m venv "$VENV"

# 4. Install deps.
"$VENV/bin/pip" install --quiet --upgrade pip
"$VENV/bin/pip" install --quiet -r "$REQS"

# 5. Verify + cache.
"$VENV/bin/python" -c "import openpyxl, formualizer"
echo "$VENV/bin/python" > "$STATE"
ok "ready — $VENV/bin/python (python $PY_VER)"
