#!/usr/bin/env bash
# install-merge-pdf-watcher.sh — instala merge_pdf_watcher como servicio systemd
# Uso: sudo ./install-merge-pdf-watcher.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
INSTALL_USER="${SUDO_USER:-mpino}"
PYTHON="$(command -v python3)"
SCRIPT="$APP_DIR/merge_pdf_watcher.py"
INBOX="/data/ias_prod/data/merge_pdf/inbox"
QUIET="30"       # segundos de quietud antes de crear batch
POLL="5"         # segundos entre comprobaciones
SERVICE="merge-pdf-watcher"
SERVICE_FILE="/etc/systemd/system/${SERVICE}.service"

echo "=== merge_pdf Inbox Watcher — Instalacion ==="
echo "  Script : $SCRIPT"
echo "  Inbox  : $INBOX"
echo "  Quietud: ${QUIET}s"
echo "  Usuario: $INSTALL_USER"
echo ""

# Asegurar que el inbox existe (el processor tambien lo crea, pero por si acaso)
mkdir -p "$INBOX"

sed \
  -e "s|{{USER}}|$INSTALL_USER|g" \
  -e "s|{{PYTHON}}|$PYTHON|g" \
  -e "s|{{SCRIPT}}|$SCRIPT|g" \
  -e "s|{{INBOX}}|$INBOX|g" \
  -e "s|{{QUIET}}|$QUIET|g" \
  -e "s|{{POLL}}|$POLL|g" \
  "$SCRIPT_DIR/merge-pdf-watcher.service.template" > "$SERVICE_FILE"

chmod 644 "$SERVICE_FILE"
systemctl daemon-reload
systemctl enable "$SERVICE"
systemctl restart "$SERVICE"

sleep 2
if systemctl is-active --quiet "$SERVICE"; then
  echo "  Servicio activo OK"
else
  echo "  ERROR: el servicio no arranco"
  journalctl -u "$SERVICE" --no-pager -n 20
  exit 1
fi

echo ""
echo "  Logs : journalctl -u $SERVICE -f"
echo "  Stop : systemctl stop $SERVICE"
echo ""
