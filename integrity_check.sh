#!/usr/bin/env bash
set -euo pipefail

ROOT="${ROOT:-$(pwd)}"

BLUEPRINT_DIR="${1:-integrity/blueprint}"
EXTRACTOR="${2:-apps/lear_cable/extractor/extract_lear_fields.py}"

RUNS_ROOT="integrity/runs"
OUT_COMPARE_ROOT="integrity/compare_out"

# extrae SCRIPT_VERSION del extractor
SCRIPT_VERSION="$(python3 - <<PY
import re
p=r"${EXTRACTOR}"
s=open(p,"r",encoding="utf-8").read()
m=re.search(r'^SCRIPT_VERSION\s*=\s*"([^"]+)"', s, re.M)
print(m.group(1) if m else "")
PY
)"
if [[ -z "$SCRIPT_VERSION" ]]; then
  echo "ERROR: cannot read SCRIPT_VERSION from ${EXTRACTOR}" >&2
  exit 2
fi

RUN_DIR="${RUNS_ROOT}/${SCRIPT_VERSION}"
OUT_DIR="${OUT_COMPARE_ROOT}/${SCRIPT_VERSION}"

echo "[INFO] BLUEPRINT: ${BLUEPRINT_DIR}"
echo "[INFO] EXTRACTOR:  ${EXTRACTOR}"
echo "[INFO] VERSION:    ${SCRIPT_VERSION}"
echo "[INFO] RUN_DIR:    ${RUN_DIR}"
echo "[INFO] OUT_DIR:    ${OUT_DIR}"
echo

# 1) Extract versionado (borra run si existe)
./run_extract_all.sh "${BLUEPRINT_DIR}" "${EXTRACTOR}" --versioned-run --runs-root "${RUNS_ROOT}"

# 2) Compare (si ya existía OUT_DIR, lo limpiamos para que sea claro)
rm -rf "${OUT_DIR}"
mkdir -p "${OUT_DIR}"

./compare_invoices_csv.sh "${BLUEPRINT_DIR}" "${RUN_DIR}" "${OUT_DIR}"

# 3) Decide pass/fail leyendo el summary
SUMMARY="${OUT_DIR}/summary.txt"

if grep -qE '^CHANGED|^MISSING' "${SUMMARY}"; then
  echo
  echo "[FAIL] Integrity check failed. See:"
  echo "  ${SUMMARY}"
  echo "  ${OUT_DIR}/diffs/"
  exit 1
fi

echo
echo "[OK] Integrity check passed for ${SCRIPT_VERSION}"
echo "  ${SUMMARY}"
exit 0