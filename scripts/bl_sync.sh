#!/bin/bash
# bl_sync.sh — Fuerza la copia de BLs desde Google Drive al inbox local
# Uso: bash bl_sync.sh
# Puede llamarse desde cron o desde la API web

set -euo pipefail

RCLONE="/usr/bin/rclone"
GDRIVE_PATH="gdrive:BLs"
LOCAL_INBOX="${BL_INBOX:-/data/ias_prod/data/bl/inbox}"
LOGFILE="/var/log/bl_sync.log"
TIMESTAMP=$(date -u '+%Y-%m-%d %H:%M:%S UTC')

RCLONE_CONF="${RCLONE_CONF:-/home/mpino/.config/rclone/rclone.conf}"

log() {
  echo "[$TIMESTAMP] $*" | tee -a "$LOGFILE"
}

log "──────────────────────────────────────"
log "Iniciando sync BL: $GDRIVE_PATH → $LOCAL_INBOX"

# Crear inbox si no existe
mkdir -p "$LOCAL_INBOX"

# Copiar (no move — sin permisos de borrado en Drive)
OUTPUT=$("$RCLONE" move "$GDRIVE_PATH" "$LOCAL_INBOX" \
  --config "$RCLONE_CONF" \
  --drive-shared-with-me \
  --include "*.pdf" \
  --include "*.PDF" \
  --stats-one-line \
  --stats 0 \
  2>&1) || true

log "$OUTPUT"
log "Sync completado"

exit 0