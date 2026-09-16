#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
TEMPLATE="${1:-$ROOT/templates/everglades.pptx}"
SOURCE="${2:-$ROOT/fixtures/sample.pptx}"
OUT="${3:-$ROOT/.tmp/rebrand}"
COLORWAY="${COLORWAY:-green}"
mkdir -p "$OUT"

if [[ -n "${PYTHON:-}" ]]; then
  PY="$PYTHON"
elif [[ -x "$ROOT/converter/.venv/bin/python" ]]; then
  PY="$ROOT/converter/.venv/bin/python"
elif [[ -x "$ROOT/converter/.venv/Scripts/python.exe" ]]; then
  PY="$ROOT/converter/.venv/Scripts/python.exe"
else
  PY="python3"
fi

"$PY" "$ROOT/rebrand/inspect_template.py" "$TEMPLATE" --out-dir "$OUT" ${BRAND_MD:+--brand-md "$BRAND_MD"} --colorway "$COLORWAY"
"$PY" "$ROOT/rebrand/inspect_source.py" "$SOURCE" --out-dir "$OUT"
"$PY" "$ROOT/rebrand/plan.py" --out-dir "$OUT" --colorway "$COLORWAY"
"$PY" "$ROOT/rebrand/build.py" --out-dir "$OUT" --source "$SOURCE" --template "$TEMPLATE" --colorway "$COLORWAY"
"$PY" "$ROOT/rebrand/qa.py" --out-dir "$OUT" --source "$SOURCE" --colorway "$COLORWAY"
echo "Wrote $OUT/OUTPUT.pptx and $OUT/QA/qa-report.md"
