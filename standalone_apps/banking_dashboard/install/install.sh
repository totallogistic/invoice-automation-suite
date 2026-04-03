#!/bin/bash
# install.sh — Banking Dashboard
# Usage:
#   sudo bash install.sh        → instala uvicorn como servicio
#   sudo bash install.sh --dev  → instala también ngrok como servicio

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(dirname "$SCRIPT_DIR")"
SERVICE_NAME="banking-dashboard"
PORT="${PORT:-8085}"
VENV="$APP_DIR/.venv"
CURRENT_USER="${SUDO_USER:-$(whoami)}"
DEV_MODE=false

[[ "$1" == "--dev" ]] && DEV_MODE=true

echo "=== Banking Dashboard installer ==="
echo "App dir : $APP_DIR"
echo "User    : $CURRENT_USER"
echo "Port    : $PORT"
echo "Dev mode: $DEV_MODE"
echo ""

# 1. Virtualenv
if [ ! -d "$VENV" ]; then
    echo "[1/5] Creating virtualenv..."
    python3 -m venv "$VENV"
else
    echo "[1/5] Virtualenv already exists."
fi

echo "[2/5] Installing Python dependencies..."
"$VENV/bin/pip" install -q --upgrade pip
"$VENV/bin/pip" install -q -r "$APP_DIR/requirements.txt"

# 2. Config files
echo "[3/5] Setting up config files..."
mkdir -p "$APP_DIR/keys" && chmod 700 "$APP_DIR/keys"
mkdir -p "$APP_DIR/data"

[ ! -f "$APP_DIR/banks.yaml" ] && cp "$APP_DIR/banks.yaml.example" "$APP_DIR/banks.yaml" && echo "      Created banks.yaml — review it!"
[ ! -f "$APP_DIR/.env" ]       && cp "$APP_DIR/.env.example" "$APP_DIR/.env"       && echo "      Created .env — fill in credentials!"

# 3. uvicorn systemd service
echo "[4/5] Installing uvicorn systemd service..."
sed \
    -e "s|{{USER}}|$CURRENT_USER|g" \
    -e "s|{{APP_DIR}}|$APP_DIR|g" \
    -e "s|{{PORT}}|$PORT|g" \
    "$SCRIPT_DIR/${SERVICE_NAME}.service.template" \
    > "/etc/systemd/system/${SERVICE_NAME}.service"

systemctl daemon-reload
systemctl enable "$SERVICE_NAME"

# 4. ngrok service (dev only)
if [ "$DEV_MODE" = true ]; then
    echo "[5/5] Installing ngrok systemd service (dev)..."
    NGROK_BIN=$(which ngrok 2>/dev/null || echo "/usr/local/bin/ngrok")
    sed \
        -e "s|{{USER}}|$CURRENT_USER|g" \
        -e "s|{{PORT}}|$PORT|g" \
        "$SCRIPT_DIR/${SERVICE_NAME}-ngrok.service.template" \
    | sed "s|/usr/local/bin/ngrok|$NGROK_BIN|g" \
        > "/etc/systemd/system/${SERVICE_NAME}-ngrok.service"
    systemctl daemon-reload
    systemctl enable "${SERVICE_NAME}-ngrok"
else
    echo "[5/5] Skipping ngrok (production mode)."
fi

echo ""
echo "✓ Done! Next steps:"
echo ""
echo "  1. Edit $APP_DIR/.env with your credentials"
echo "  2. sudo systemctl start $SERVICE_NAME"
if [ "$DEV_MODE" = true ]; then
echo "  3. sudo systemctl start ${SERVICE_NAME}-ngrok"
echo "  4. Check ngrok URL: journalctl -u ${SERVICE_NAME}-ngrok -f"
echo ""
echo "  ⚠️  After ngrok starts, update the redirect URL in Enable Banking dashboard"
fi
echo ""
echo "  Logs: journalctl -u $SERVICE_NAME -f"