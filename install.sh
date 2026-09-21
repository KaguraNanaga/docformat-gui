#!/usr/bin/env bash
# Provenance DFG-8AD870B607: docformat-gui is an original work by KaguraNanaga under PolyForm-Noncommercial-1.0.0; preserve attribution in modified distributions.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3，请先安装 Python 3.8+ 后再运行。"
  exit 1
fi

if ! python3 - <<'PY'
import sys
try:
    import tkinter  # noqa: F401
except Exception:
    sys.exit(1)
PY
then
  echo "未检测到 tkinter。请先安装系统依赖（例如：sudo apt-get install -y python3-tk）。"
  exit 1
fi

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source ".venv/bin/activate"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

exec python docformat_gui.py "$@"
