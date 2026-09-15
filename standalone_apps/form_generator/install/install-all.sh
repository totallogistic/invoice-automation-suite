#!/bin/bash
#
# Form Generator - Multi-Environment Installation
# Installs one service per .env.* file found
#

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

# Banner
echo -e "${BLUE}"
echo "============================================================"
echo "  📋 Form Generator - Multi-Environment Installer"
echo "============================================================"
echo -e "${NC}"

# Check root
if [ "$EUID" -ne 0 ]; then
  echo -e "${RED}❌ Please run as root (use sudo)${NC}"
  exit 1
fi

# Detect paths
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(dirname "$SCRIPT_DIR")"
REPO_ROOT="$(cd "$APP_DIR/../.." && pwd)"

echo -e "${BLUE}🔍 Scanning for environment files in:${NC}"
echo "   $REPO_ROOT"
echo ""

# Find all .env.* files
ENV_FILES=$(find "$REPO_ROOT" -maxdepth 1 -name ".env.*" -type f | sort)

if [ -z "$ENV_FILES" ]; then
  echo -e "${YELLOW}⚠️  No .env.* files found${NC}"
  echo "   Looking for: .env.prod, .env.dev, etc."
  exit 1
fi

# Count environments
ENV_COUNT=$(echo "$ENV_FILES" | wc -l)
echo -e "${GREEN}✓ Found $ENV_COUNT environment(s):${NC}"
while IFS= read -r file; do
  basename "$file"
done <<<"$ENV_FILES"
echo ""

# Confirm
read -p "Install Form Generator for all environments? [Y/n] " -n 1 -r
echo
if [[ $REPLY =~ ^[Nn]$ ]]; then
  echo "Cancelled."
  exit 0
fi

# ── Dependencias Python (una vez, en el --user site del usuario del servicio) ──
# El servicio corre como $SUDO_USER con /usr/bin/python3 (sin venv), así que las
# deps deben vivir en SU --user site. Instalar aquí evita el ModuleNotFoundError
# al añadir paquetes al requirements.txt (fix deploy §7-D). Instalación desatendida.
REQ_FILE="$APP_DIR/requirements.txt"
if [ -f "$REQ_FILE" ]; then
  echo -e "${BLUE}📦 Instalando dependencias Python desde requirements.txt...${NC}"
  if [ -n "$SUDO_USER" ]; then
    sudo -H -u "$SUDO_USER" python3 -m pip install --user --break-system-packages -r "$REQ_FILE"
  else
    python3 -m pip install --break-system-packages -r "$REQ_FILE"
  fi
  echo -e "${GREEN}   ✓ Dependencias instaladas${NC}"
  echo ""
else
  echo -e "${YELLOW}⚠️  No hay requirements.txt en $APP_DIR, se omite la instalación de deps${NC}"
  echo ""
fi

# Install each environment - AVOID SUBSHELL with <<< syntax
INSTALL_COUNT=0
echo ""

while IFS= read -r ENV_FILE; do
  ENV_NAME=$(basename "$ENV_FILE" | sed 's/^\.env\.//')

  echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
  echo -e "${BLUE}📦 Installing: ${GREEN}$ENV_NAME${NC}"
  echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
  echo ""

  # Read only FORM_UI_PORT to check if this env file is meant for the form generator
  FORM_PORT=$(grep "^FORM_UI_PORT=" "$ENV_FILE" 2>/dev/null | head -1 | cut -d'=' -f2 | tr -d ' "'"'"'')
  SMTP_HOST=$(grep "^SMTP_HOST=" "$ENV_FILE" 2>/dev/null | head -1 | cut -d'=' -f2 | tr -d ' "'"'"'')

  if [ -z "$FORM_PORT" ]; then
    echo -e "${YELLOW}⚠️  No FORM_UI_PORT in $ENV_FILE, skipping${NC}"
    echo ""
    continue
  fi

  echo -e "${BLUE}   Environment: ${GREEN}$ENV_NAME${NC}"
  echo -e "${BLUE}   Port:        ${GREEN}$FORM_PORT${NC}"
  echo -e "${BLUE}   Env file:    ${GREEN}$ENV_FILE${NC}"
  if [ -n "$SMTP_HOST" ]; then
    echo -e "${BLUE}   Email:       ${GREEN}Configured (all MAIL_TO_* vars loaded from env file)${NC}"
  else
    echo -e "${BLUE}   Email:       ${YELLOW}Not configured${NC}"
  fi
  echo ""

  # Create service name
  SERVICE_NAME="form-generator-${ENV_NAME}"
  SERVICE_FILE="/etc/systemd/system/${SERVICE_NAME}.service"

  # Check if port already in use
  if lsof -Pi :$FORM_PORT -sTCP:LISTEN -t >/dev/null 2>&1; then
    echo -e "${YELLOW}⚠️  Port $FORM_PORT already in use, skipping${NC}"
    echo ""
    continue
  fi

  # Generate service file with EnvironmentFile to load ALL variables from env file
  cat >"$SERVICE_FILE" <<SERVICEEOF
[Unit]
Description=JSON Schema Form Generator [$ENV_NAME]
After=network.target
Documentation=https://github.com/totallogistic/invoice-automation-suite

[Service]
Type=simple
User=$SUDO_USER
WorkingDirectory=$APP_DIR
ExecStart=$(command -v python3) $APP_DIR/app.py

# Restart policy
Restart=always
RestartSec=10

# Logging
StandardOutput=journal
StandardError=journal
SyslogIdentifier=${SERVICE_NAME}

# Environment - load ALL variables from env file (MAIL_TO_*, SMTP_*, etc.)
Environment="PYTHONUNBUFFERED=1"
Environment="PYTHONDONTWRITEBYTECODE=1"
EnvironmentFile=$ENV_FILE

# Security
NoNewPrivileges=true
PrivateTmp=true

[Install]
WantedBy=multi-user.target
SERVICEEOF

  chmod 644 "$SERVICE_FILE"

  # Reload and enable
  systemctl daemon-reload
  systemctl enable "$SERVICE_NAME" >/dev/null 2>&1

  # Start service
  if systemctl start "$SERVICE_NAME"; then
    sleep 2
    if systemctl is-active --quiet "$SERVICE_NAME"; then
      echo -e "${GREEN}   ✓ Service started successfully${NC}"
      INSTALL_COUNT=$((INSTALL_COUNT + 1))
    else
      echo -e "${RED}   ✗ Service failed to start${NC}"
      echo -e "${YELLOW}   Check: sudo journalctl -u $SERVICE_NAME -n 20${NC}"
    fi
  else
    echo -e "${RED}   ✗ Failed to start service${NC}"
  fi

  echo ""
done <<<"$ENV_FILES"

# Summary
echo -e "${GREEN}============================================================${NC}"
echo -e "${GREEN}  ✓ Installation Complete!${NC}"
echo -e "${GREEN}============================================================${NC}"
echo ""
echo -e "${BLUE}📊 Installed Services:${NC}"
systemctl list-units --type=service --all | grep form-generator
echo ""
echo -e "${BLUE}🔧 Management Commands:${NC}"
echo "   View all:    sudo systemctl status 'form-generator-*'"
echo "   Stop all:    sudo systemctl stop 'form-generator-*'"
echo "   Start all:   sudo systemctl start 'form-generator-*'"
echo "   Logs (prod): sudo journalctl -u form-generator-prod -f"
echo "   Logs (dev):  sudo journalctl -u form-generator-dev -f"
echo ""
echo -e "${BLUE}🌐 Access:${NC}"

# Show access URLs - ALSO AVOID SUBSHELL
while IFS= read -r ENV_FILE; do
  ENV_NAME=$(basename "$ENV_FILE" | sed 's/^\.env\.//')
  FORM_PORT=$(grep "^FORM_UI_PORT=" "$ENV_FILE" 2>/dev/null | head -1 | cut -d'=' -f2 | tr -d ' "'"'"'')
  if [ -n "$FORM_PORT" ]; then
    HOST_IP=$(hostname -I | awk '{print $1}')
    echo "   $ENV_NAME: http://$HOST_IP:$FORM_PORT"
  fi
done <<<"$ENV_FILES"

echo ""
