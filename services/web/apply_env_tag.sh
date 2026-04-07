#!/usr/bin/env bash
set -euo pipefail

ENV="${1:-}"
SITE_DIR="${2:-}"

if [[ -z "${ENV}" || -z "${SITE_DIR}" ]]; then
  echo "Usage: $0 <ENV> <SITE_DIR>" >&2
  exit 2
fi

patch_file() {
  local f="$1"
  [[ -f "$f" ]] || return 0

  # TITLE
  if grep -qE '<title>\[[^]]+\]' "$f"; then
    sed -i '' -E "s#<title>\\[[^]]+\\][[:space:]]*#<title>[${ENV}] #g" "$f"
  else
    sed -i '' -E "s#<title>#<title>[${ENV}] #g" "$f"
  fi

  # H1
  if grep -qE '<h1>\[[^]]+\]' "$f"; then
    sed -i '' -E "s#<h1>\\[[^]]+\\][[:space:]]*#<h1>[${ENV}] #g" "$f"
  else
    sed -i '' -E "s#<h1>#<h1>[${ENV}] #g" "$f"
  fi
}

# landing
patch_file "${SITE_DIR}/index.html"

# todas las tools
if [[ -d "${SITE_DIR}/tools" ]]; then
  while IFS= read -r f; do
    patch_file "$f"
  done < <(find "${SITE_DIR}/tools" -mindepth 2 -maxdepth 2 -type f -name index.html | sort)
fi
