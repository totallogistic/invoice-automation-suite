#!/usr/bin/env bash
set -euo pipefail
shopt -s nullglob

usage() {
  cat >&2 <<'EOF'
Usage:
  run_extract_all.sh <root_folder> [path_to_extract_lear_fields.py] [--merge] [--merge-only] [--merge-out <file_base>]
                     [--versioned-run] [--runs-root <dir>] [--no-clean]

Examples:
  # 1) Extract en cada subfolder (1 nivel), outputs en el MISMO subfolder (modo actual)
  ./run_extract_all.sh /ruta/a/root

  # 2) Extract + merge (1 item por lote a partir de invoices_extracted_summary.json)
  ./run_extract_all.sh /ruta/a/root --merge

  # 3) Solo merge (no re-extrae nada)
  ./run_extract_all.sh /ruta/a/root --merge-only

  # 4) Extract + merge con output base custom (crea .json y .csv)
  ./run_extract_all.sh /ruta/a/root --merge --merge-out /ruta/a/root/merged/lotes

  # 5) Integrity run versionado: inputs desde blueprint, outputs en integrity/runs/<SCRIPT_VERSION>/<lote>/
  ./run_extract_all.sh integrity/blueprint apps/lear_cable/extractor/extract_lear_fields.py --versioned-run

Notes:
  - El merge genera: <file_base>.json y <file_base>.csv
  - Si no das --merge-out, usa: <root_efectivo>/merged_lot_summaries
  - Con --versioned-run:
      - root_folder es el blueprint (inputs)
      - outputs van a runs-root/SCRIPT_VERSION/<lote>/
      - si ya existe esa carpeta, se borra salvo que uses --no-clean
EOF
}

if [[ $# -lt 1 ]]; then usage; exit 2; fi

ROOT_DIR="$1"; shift

PY_EXTRACT="extract_lear_fields.py"
DO_MERGE=0
MERGE_ONLY=0
MERGE_BASE=""

VERSIONED_RUN=0
RUNS_ROOT="integrity/runs"
CLEAN_EXISTING=1

# Segundo argumento opcional: ruta al extractor (si no empieza por -)
if [[ $# -gt 0 && "${1:0:1}" != "-" ]]; then
  PY_EXTRACT="$1"
  shift
fi

# Flags
while [[ $# -gt 0 ]]; do
  case "$1" in
    --merge)
      DO_MERGE=1
      shift
      ;;
    --merge-only)
      DO_MERGE=1
      MERGE_ONLY=1
      shift
      ;;
    --merge-out)
      [[ $# -ge 2 ]] || { echo "ERROR: --merge-out requiere un valor" >&2; exit 2; }
      MERGE_BASE="$2"
      shift 2
      ;;
    --versioned-run)
      VERSIONED_RUN=1
      shift
      ;;
    --runs-root)
      [[ $# -ge 2 ]] || { echo "ERROR: --runs-root requiere un valor" >&2; exit 2; }
      RUNS_ROOT="$2"
      shift 2
      ;;
    --no-clean)
      CLEAN_EXISTING=0
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "ERROR: argumento desconocido: $1" >&2
      usage
      exit 2
      ;;
  esac
done

if [[ ! -d "${ROOT_DIR}" ]]; then
  echo "ERROR: Root folder not found: ${ROOT_DIR}" >&2
  exit 2
fi

if (( MERGE_ONLY == 0 )); then
  if [[ ! -f "${PY_EXTRACT}" ]]; then
    echo "ERROR: Python extractor not found: ${PY_EXTRACT}" >&2
    exit 2
  fi
fi

# Resolve SCRIPT_VERSION if needed
SCRIPT_VERSION=""
RUN_ROOT=""

if (( VERSIONED_RUN == 1 )); then
  if (( MERGE_ONLY == 1 )); then
    echo "ERROR: --versioned-run con --merge-only no tiene sentido (no hay outputs a comparar/generar)." >&2
    exit 2
  fi

  SCRIPT_VERSION="$(python3 - <<PY
import re, sys
p = r"${PY_EXTRACT}"
s = open(p, "r", encoding="utf-8").read()
m = re.search(r'^SCRIPT_VERSION\\s*=\\s*"([^"]+)"', s, re.M)
print(m.group(1) if m else "")
PY
)"
  if [[ -z "${SCRIPT_VERSION}" ]]; then
    echo "ERROR: Could not read SCRIPT_VERSION from: ${PY_EXTRACT}" >&2
    exit 2
  fi

  RUN_ROOT="${RUNS_ROOT%/}/${SCRIPT_VERSION}"
  if [[ -d "${RUN_ROOT}" && "${CLEAN_EXISTING}" -eq 1 ]]; then
    echo "[INFO] Removing existing run folder (same version): ${RUN_ROOT}"
    rm -rf "${RUN_ROOT}"
  fi
  mkdir -p "${RUN_ROOT}"
fi

# Effective root for merge defaults
EFFECTIVE_ROOT="${ROOT_DIR%/}"
if (( VERSIONED_RUN == 1 )); then
  EFFECTIVE_ROOT="${RUN_ROOT}"
fi

if [[ -z "${MERGE_BASE}" ]]; then
  MERGE_BASE="${EFFECTIVE_ROOT%/}/merged_lot_summaries"
fi

echo "[INFO] Input root: ${ROOT_DIR}"
echo "[INFO] Extractor: ${PY_EXTRACT}"
echo "[INFO] Versioned-run: ${VERSIONED_RUN}"
if (( VERSIONED_RUN == 1 )); then
  echo "[INFO] SCRIPT_VERSION: ${SCRIPT_VERSION}"
  echo "[INFO] Runs root: ${RUNS_ROOT}"
  echo "[INFO] Run folder: ${RUN_ROOT}"
fi
echo "[INFO] Merge: ${DO_MERGE}"
echo "[INFO] Merge-only: ${MERGE_ONLY}"
echo "[INFO] Merge base: ${MERGE_BASE}"
echo

# 1) Extract por subcarpeta (1 nivel)
if (( MERGE_ONLY == 0 )); then
  for dir in "${ROOT_DIR%/}"/*/; do
    [[ -d "${dir}" ]] || continue

    pdfs=( "${dir}"*.pdf "${dir}"*.PDF )
    if (( ${#pdfs[@]} == 0 )); then
      echo "[SKIP] No PDFs in: ${dir}"
      continue
    fi

    if (( VERSIONED_RUN == 1 )); then
      lot="$(basename "${dir%/}")"
      out_dir="${RUN_ROOT%/}/${lot}/"
      mkdir -p "${out_dir}"

      echo "[RUN ] ${dir} -> ${out_dir} (pdfs=${#pdfs[@]})"
      python3 "${PY_EXTRACT}" "${dir}" -o "${out_dir}"
      echo "[ OK ] ${lot}"
      echo
    else
      echo "[RUN ] ${dir} (pdfs=${#pdfs[@]})"
      # outputs en el MISMO subfolder
      python3 "${PY_EXTRACT}" "${dir}" -o "${dir}"
      echo "[ OK ] ${dir}"
      echo
    fi
  done
else
  echo "[INFO] --merge-only: skipping extraction step."
  echo
fi

# 2) Merge opcional: 1 item/fila por lote
if (( DO_MERGE == 1 )); then
  echo "[MERGE] Building merged summaries…"

  python3 - <<PY
import csv, json
from pathlib import Path

root = Path(r"${EFFECTIVE_ROOT}").resolve()
merge_base = Path(r"${MERGE_BASE}").resolve()
merge_base.parent.mkdir(parents=True, exist_ok=True)

rows = []
for d in sorted([p for p in root.iterdir() if p.is_dir()]):
    summary = d / "invoices_extracted_summary.json"
    if not summary.exists():
        continue
    try:
        data = json.loads(summary.read_text(encoding="utf-8"))
    except Exception:
        continue

    rows.append({
        "lot": d.name,
        "path": str(d),
        "count_files": data.get("count_files"),
        "total_gross_weight": data.get("total_gross_weight"),
        "total_total_invoice": data.get("total_total_invoice"),
    })

out_json = merge_base.with_suffix(".json")
out_csv  = merge_base.with_suffix(".csv")

out_json.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")

with out_csv.open("w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(
        f,
        fieldnames=["lot","path","count_files","total_gross_weight","total_total_invoice"]
    )
    w.writeheader()
    w.writerows(rows)

print(f"[ OK ] merged → {out_json}")
print(f"[ OK ] merged → {out_csv}")
print(f"[INFO] lots merged: {len(rows)}")
PY

  echo "[DONE] Merge complete."
fi

echo "[DONE]"