#!/bin/bash
# install.sh — Banking Dashboard
# Usage: sudo bash install.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(dirname "$SCRIPT_DIR")"
REPO_DIR="$(dirname "$(dirname "$APP_DIR")")"
SERVICE_NAME="banking-dashboard"
PORT="${PORT:-8080}"
VENV="$APP_DIR/.venv"
CURRENT_USER="${SUDO_USER:-$(whoami)}"

echo "=== Banking Dashboard installer ==="
echo "App dir : $APP_DIR"
echo "Repo dir: $REPO_DIR"
echo "User    : $CURRENT_USER"
echo "Port    : $PORT"
echo ""

# 1. Virtual environment
if [ ! -d "$VENV" ]; then
  echo "[1/5] Creating virtualenv..."
  python3 -m venv "$VENV"
fi

echo "[2/5] Installing dependencies..."
"$VENV/bin/pip" install -q --upgrade pip
"$VENV/bin/pip" install -q -r "$APP_DIR/requirements.txt"

# 2. Keys directory
mkdir -p "$APP_DIR/keys"
chmod 700 "$APP_DIR/keys"

# 3. Config files
if [ ! -f "$APP_DIR/banks.yaml" ]; then
  cp "$APP_DIR/banks.yaml.example" "$APP_DIR/banks.yaml"
  echo "[3/5] Created banks.yaml from example — review it before starting!"
else
  echo "[3/5] banks.yaml already exists — skipping."
fi

if [ ! -f "$APP_DIR/.env" ]; then
  cp "$APP_DIR/.env.example" "$APP_DIR/.env"
  echo "      Created .env from example — FILL IN your Enable Banking credentials!"
else
  echo "      .env already exists — skipping."
fi

# 4. systemd service
echo "[4/5] Installing systemd service..."
SERVICE_FILE="/etc/systemd/system/${SERVICE_NAME}.service"
sed \
  -e "s|{{USER}}|$CURRENT_USER|g" \
  -e "s|{{INSTALL_DIR}}|$REPO_DIR|g" \
  -e "s|{{VENV}}|$VENV|g" \
  -e "s|{{PORT}}|$PORT|g" \
  "$SCRIPT_DIR/${SERVICE_NAME}.service.template" > "$SERVICE_FILE"

systemctl daemon-reload
systemctl enable "$SERVICE_NAME"

# 5. Done
echo "[5/5] Done!"
echo ""
echo "Next steps:"
echo "  1. Edit $APP_DIR/.env — add ENABLE_BANKING_APP_ID and KEY_PATH"
echo "  2. Generate RSA key pair (if you don't have one):"
echo "       openssl genrsa -out $APP_DIR/keys/private.pem 2048"
echo "       openssl rsa -in $APP_DIR/keys/private.pem -pubout -out $APP_DIR/keys/public.pem"
echo "     Then upload the public.pem to enablebanking.com dashboard."
echo "  3. Register callback URL on Enable Banking: http://YOUR_HOST:$PORT/callback"
echo "  4. sudo systemctl start $SERVICE_NAME"
echo "  5. Open http://localhost:$PORT/setup and connect each bank"
