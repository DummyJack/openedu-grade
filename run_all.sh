#!/usr/bin/env bash
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

if [[ -x "$SCRIPT_DIR/venv/bin/python" ]]; then
  PYTHON_CMD="$SCRIPT_DIR/venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_CMD="python"
else
  echo "[ERROR] Python not found."
  exit 1
fi

run_step() {
  local script_name="$1"
  local retry_hint="$2"

  echo ""
  echo "=================================================="
  echo "[RUN] $script_name"
  echo "=================================================="

  "$PYTHON_CMD" "$script_name"
  local exit_code=$?

  if [[ $exit_code -ne 0 ]]; then
    echo ""
    echo "[FAIL] $script_name failed (exit code: $exit_code)"
    echo "[RETRY] $retry_hint"
    exit $exit_code
  fi
}

echo "[INFO] Using Python: $PYTHON_CMD"
echo "[INFO] Project Dir: $SCRIPT_DIR"
echo "[INFO] Cleaning old output files..."

mkdir -p "$SCRIPT_DIR/scores" "$SCRIPT_DIR/errors"
rm -f "$SCRIPT_DIR/openedu_all_videos.xlsx"
rm -f "$SCRIPT_DIR/scores/"*.xlsx
rm -f "$SCRIPT_DIR/errors/"*.xlsx

run_step "student.py" \
  "Please rerun student.py. If some users failed, check errors/video_errors.xlsx."

if [[ ! -f "$SCRIPT_DIR/openedu_all_videos.xlsx" ]]; then
  echo ""
  echo "[FAIL] openedu_all_videos.xlsx was not generated."
  echo "[RETRY] Please rerun student.py first."
  exit 1
fi

run_step "video_grade.py" \
  "Please rerun video_grade.py (it will ask chapter range again)."

run_step "grade.py" \
  "Please rerun grade.py. If partial failures exist, choose option 3 to reprocess errors/grade_errors.xlsx."

echo ""
echo "[DONE] All scripts completed successfully."
