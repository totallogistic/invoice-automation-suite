#!/usr/bin/env bash
# install-dua-watcher.sh — instala dua_watcher como servicio systemd.
# Uso: sudo ./install-dua-watcher.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
INSTALL_USER="${SUDO_USER:-mpino}"
PYTHON="$(command -v python3)"
SCRIPT="$APP_DIR/dua_watcher.py"
BASE="/data/ias_prod/data/dua"
INBOX="$BASE/inbox"
QUIET="30"       # segundos de quietud antes de procesar
POLL="5"
SERVICE="dua-watcher"
SERVICE_FILE="/etc/systemd/system/${SERVICE}.service"

echo "=== DUA Inbox Watcher — Instalacion ==="
echo "  Script : $SCRIPT"
echo "  Inbox  : $INBOX"
echo "  Usuario: $INSTALL_USER"
echo ""

# Prerequisitos de sistema para OCR
echo "[1/4] Comprobando dependencias de sistema (tesseract, poppler)..."
MISSING=""
command -v tesseract >/dev/null 2>&1 || MISSING="$MISSING tesseract-ocr"
command -v pdftoppm  >/dev/null 2>&1 || MISSING="$MISSING poppler-utils"
if [ -n "$MISSING" ]; then
  echo "    Instalando:$MISSING"
  apt-get update -qq && apt-get install -y $MISSING
else
  echo "    OK"
fi

# Dependencias Python
echo "[2/4] Instalando dependencias Python..."
sudo -H -u "$INSTALL_USER" python3 -m pip install --user --break-system-packages \
  -r "$APP_DIR/requirements.txt" >/dev/null 2>&1 || \
  python3 -m pip install --break-system-packages -r "$APP_DIR/requirements.txt"
echo "    OK"

# Carpetas de trabajo
echo "[3/4] Creando carpetas en $BASE ..."
mkdir -p "$BASE"/{inbox,outbox,review,processed,error}
chown -R "$INSTALL_USER" "$BASE" || true
echo "    OK"

# Fichero de entorno (SMTP / destinatario revision) si no existe
ENVFILE="$APP_DIR/dua.env"
if [ ! -f "$ENVFILE" ]; then
  cat > "$ENVFILE" <<'EOF'
# Configuracion del dua_watcher (SMTP para emails de revision).
# Reutiliza el mismo relay que el resto del stack.
SMTP_HOST=172.18.0.1
SMTP_PORT=25
# SMTP_USER=
# SMTP_PASS=
MAIL_FROM=procesos@totallogistic.es
# Destinatario(s) interno(s) que revisan los DUA dudosos (coma-separado):
MAIL_INTERNO_DUA=miguel.pino@codeengtools.eu
EOF
  echo "    Creado $ENVFILE (revisa MAIL_INTERNO_DUA y SMTP)"
fi

# Servicio systemd
echo "[4/4] Instalando servicio systemd..."
sed \
  -e "s|{{USER}}|$INSTALL_USER|g" \
  -e "s|{{PYTHON}}|$PYTHON|g" \
  -e "s|{{SCRIPT}}|$SCRIPT|g" \
  -e "s|{{APPDIR}}|$APP_DIR|g" \
  -e "s|{{INBOX}}|$INBOX|g" \
  -e "s|{{QUIET}}|$QUIET|g" \
  -e "s|{{POLL}}|$POLL|g" \
  "$SCRIPT_DIR/dua-watcher.service.template" > "$SERVICE_FILE"
chmod 644 "$SERVICE_FILE"
systemctl daemon-reload
systemctl enable "$SERVICE"
systemctl restart "$SERVICE"

sleep 2
if systemctl is-active --quiet "$SERVICE"; then
  echo "    Servicio activo OK"
else
  echo "    ERROR: el servicio no arranco"
  journalctl -u "$SERVICE" --no-pager -n 20
  exit 1
fi

echo ""
echo "  Logs : journalctl -u $SERVICE -f"
echo "  Env  : $ENVFILE  (SMTP + MAIL_INTERNO_DUA)"
echo "  Carpetas: $BASE/{inbox,outbox,review,processed,error}"
echo ""
echo "  Siguiente: comparte inbox/ y outbox/ por Samba con install-samba-dua.sh"
echo ""
