#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

DEFAULT_PYTHON="/Users/kumagaihitoshi/anaconda3/envs/pseudo-semantic-bridge/bin/python"
if [[ -n "${PYTHON_BIN:-}" ]]; then
  PYTHON="$PYTHON_BIN"
elif [[ -x "$DEFAULT_PYTHON" ]]; then
  PYTHON="$DEFAULT_PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON="$(command -v python3)"
else
  PYTHON="$(command -v python)"
fi

OUT_DIR=".tracecov"
SUMMARY_FILE=".trace-summary-missing.txt"

rm -rf "$OUT_DIR" "$SUMMARY_FILE"

"$PYTHON" -m trace \
  --count \
  --missing \
  --summary \
  --coverdir "$OUT_DIR" \
  --module pytest -- -q "$@" > "$SUMMARY_FILE"

echo "== src coverage by module =="
grep "src\\." "$SUMMARY_FILE" || true

echo
awk '
  /src\./ {
    pct = $2
    sub(/%/, "", pct)
    lines += $1
    covered += ($1 * pct / 100)
  }
  END {
    if (lines > 0) {
      printf("SRC_TOTAL_LINES=%d\n", lines)
      printf("SRC_WEIGHTED_COVERAGE=%.2f%%\n", 100 * covered / lines)
    } else {
      print("SRC_TOTAL_LINES=0")
      print("SRC_WEIGHTED_COVERAGE=0.00%")
    }
  }
' "$SUMMARY_FILE"

echo
echo "Saved:"
echo "  $OUT_DIR/"
echo "  $SUMMARY_FILE"
