#!/usr/bin/env bash
# install-bl-watcher.sh — instala bl_inbox_watcher como servicio systemd
# Uso: sudo ./install-bl-watcher.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INSTALL_USER="${SUDO_USER:-mpino}"
PYTHON="$(command -v python3)"
SCRIPT="$SCRIPT_DIR/bl_inbox_watcher.py"
INBOX="/data/ias_prod/data/bl/inbox"
QUIET="15"      # segundos de quietud antes de crear batch
POLL="3"        # segundos entre comprobaciones
SERVICE="bl-inbox-watcher"
SERVICE_FILE="/etc/systemd/system/${SERVICE}.service"

echo "=== BL Inbox Watcher — Instalación ==="
echo "  Script : $SCRIPT"
echo "  Inbox  : $INBOX"
echo "  Quietud: ${QUIET}s"
echo "  Usuario: $INSTALL_USER"
echo ""

sed \
  -e "s|{{USER}}|$INSTALL_USER|g" \
  -e "s|{{PYTHON}}|$PYTHON|g" \
  -e "s|{{SCRIPT}}|$SCRIPT|g" \
  -e "s|{{INBOX}}|$INBOX|g" \
  -e "s|{{QUIET}}|$QUIET|g" \
  -e "s|{{POLL}}|$POLL|g" \
  "$SCRIPT_DIR/bl-inbox-watcher.service.template" > "$SERVICE_FILE"

chmod 644 "$SERVICE_FILE"
systemctl daemon-reload
systemctl enable "$SERVICE"
systemctl restart "$SERVICE"

sleep 2
if systemctl is-active --quiet "$SERVICE"; then
  echo "  Servicio activo ✓"
else
  echo "  ERROR: el servicio no arrancó"
  journalctl -u "$SERVICE" --no-pager -n 20
  exit 1
fi

echo ""
echo "  Logs : journalctl -u $SERVICE -f"
echo "  Stop : systemctl stop $SERVICE"
echo ""