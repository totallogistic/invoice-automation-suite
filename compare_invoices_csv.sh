#!/usr/bin/env bash
set -euo pipefail

# compare_invoices_csv.sh (bash 3.2 compatible)
# Usage:
#   ./compare_invoices_csv.sh /path/to/runA /path/to/runB [/path/to/outdir]

RUN_A="${1:-}"
RUN_B="${2:-}"
OUTDIR="${3:-./csv_diff_out}"

if [[ -z "$RUN_A" || -z "$RUN_B" ]]; then
  echo "Usage: $0 <runA_root> <runB_root> [outdir]" >&2
  exit 2
fi
if [[ ! -d "$RUN_A" || ! -d "$RUN_B" ]]; then
  echo "Error: both run roots must be directories" >&2
  exit 2
fi

mkdir -p "$OUTDIR/diffs" "$OUTDIR/normalized/A" "$OUTDIR/normalized/B"

TMP_A="$(mktemp)"
TMP_B="$(mktemp)"
TMP_KEYS="$(mktemp)"

# relative paths like ./batch1/invoices_extracted.csv
( cd "$RUN_A" && find . -type f -name 'invoices_extracted.csv' -print | LC_ALL=C sort ) > "$TMP_A"
( cd "$RUN_B" && find . -type f -name 'invoices_extracted.csv' -print | LC_ALL=C sort ) > "$TMP_B"
cat "$TMP_A" "$TMP_B" | LC_ALL=C sort -u > "$TMP_KEYS"

normalize_csv() {
  src="$1"
  dst="$2"

  lines="$(wc -l < "$src" | tr -d ' ')"
  if [[ "$lines" -le 1 ]]; then
    cp -f "$src" "$dst"
    return
  fi

  {
    head -n 1 "$src"
    tail -n +2 "$src" | LC_ALL=C sort
  } > "$dst"
}

# helper: check if rel path exists in list file (sorted)
exists_in_list() {
  rel="$1"
  listfile="$2"
  grep -Fqx "$rel" "$listfile"
}

SUMMARY="$OUTDIR/summary.txt"
: > "$SUMMARY"

ok=0
changed=0
missing=0

while IFS= read -r rel; do
  a_exists=0
  b_exists=0
  if exists_in_list "$rel" "$TMP_A"; then a_exists=1; fi
  if exists_in_list "$rel" "$TMP_B"; then b_exists=1; fi

  if [[ "$a_exists" -eq 0 ]]; then
    echo "MISSING_IN_A  $rel" | tee -a "$SUMMARY" >/dev/null
    missing=$((missing+1))
    continue
  fi
  if [[ "$b_exists" -eq 0 ]]; then
    echo "MISSING_IN_B  $rel" | tee -a "$SUMMARY" >/dev/null
    missing=$((missing+1))
    continue
  fi

  a_path="$RUN_A/$rel"
  b_path="$RUN_B/$rel"

  norm_a="$OUTDIR/normalized/A/${rel#./}"
  norm_b="$OUTDIR/normalized/B/${rel#./}"
  mkdir -p "$(dirname "$norm_a")" "$(dirname "$norm_b")"

  normalize_csv "$a_path" "$norm_a"
  normalize_csv "$b_path" "$norm_b"

  if diff -q "$norm_a" "$norm_b" >/dev/null 2>&1; then
    echo "OK           $rel" | tee -a "$SUMMARY" >/dev/null
    ok=$((ok+1))
  else
    echo "CHANGED      $rel" | tee -a "$SUMMARY" >/dev/null
    changed=$((changed+1))

    diff_out="$OUTDIR/diffs/${rel#./}.diff"
    mkdir -p "$(dirname "$diff_out")"
    diff -u "$norm_a" "$norm_b" > "$diff_out" || true
  fi
done < "$TMP_KEYS"

rm -f "$TMP_A" "$TMP_B" "$TMP_KEYS"

echo
echo "==== RESULT ===="
echo "OK:      $ok"
echo "CHANGED: $changed"
echo "MISSING: $missing"
echo
echo "Summary: $SUMMARY"
echo "Diffs:   $OUTDIR/diffs/"