#!/bin/bash
#
# Form Generator - Installation Script
# Installs the form generator as a systemd service
#
# Usage:
#   cd standalone_apps/form_generator
#   sudo ./install/install.sh
#

set -e # Exit on error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Banner
echo -e "${BLUE}"
echo "============================================================"
echo "  📋 Form Generator - Installation Script"
echo "============================================================"
echo -e "${NC}"

# Check if running as root
if [ "$EUID" -ne 0 ]; then
  echo -e "${RED}❌ Please run as root (use sudo)${NC}"
  exit 1
fi

# Detect paths automatically
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(dirname "$SCRIPT_DIR")"
REPO_ROOT="$(cd "$APP_DIR/../.." && pwd)"

echo -e "${BLUE}📁 Detected paths:${NC}"
echo "   Repo root: $REPO_ROOT"
echo "   App dir:   $APP_DIR"
echo ""

# Detect current user (the one who called sudo)
INSTALL_USER="${SUDO_USER:-$USER}"
if [ "$INSTALL_USER" = "root" ]; then
  echo -e "${RED}❌ Cannot detect non-root user. Please run with 'sudo' as a regular user.${NC}"
  exit 1
fi

echo -e "${BLUE}👤 Installing for user: ${GREEN}$INSTALL_USER${NC}"
echo ""

# Check if app.py exists
if [ ! -f "$APP_DIR/app.py" ]; then
  echo -e "${RED}❌ app.py not found in $APP_DIR${NC}"
  exit 1
fi

# Check Python version
echo -e "${BLUE}🐍 Checking Python...${NC}"
PYTHON_CMD=$(command -v python3)
if [ -z "$PYTHON_CMD" ]; then
  echo -e "${RED}❌ python3 not found. Please install Python 3.${NC}"
  exit 1
fi

PYTHON_VERSION=$($PYTHON_CMD --version | awk '{print $2}')
echo -e "   Found: $PYTHON_CMD (version $PYTHON_VERSION)"

# Check if requirements.txt exists
if [ ! -f "$APP_DIR/requirements.txt" ]; then
  echo -e "${RED}❌ requirements.txt not found${NC}"
  exit 1
fi

# Ask if user wants to install dependencies
echo ""
echo -e "${YELLOW}📦 Install Python dependencies?${NC}"
read -p "   This will run: pip install -r requirements.txt [Y/n] " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Nn]$ ]]; then
  echo -e "${BLUE}   Installing dependencies...${NC}"

  # Use --break-system-packages for Raspberry Pi OS (PEP 668)
  sudo -u $INSTALL_USER pip install --user --break-system-packages -r "$APP_DIR/requirements.txt"

  echo -e "${GREEN}   ✓ Dependencies installed${NC}"
else
  echo -e "${YELLOW}   ⚠ Skipped dependency installation${NC}"
fi

# Create systemd service file
echo ""
echo -e "${BLUE}⚙️  Creating systemd service...${NC}"

SERVICE_FILE="/etc/systemd/system/form-generator.service"
TEMPLATE_FILE="$SCRIPT_DIR/form-generator.service.template"

if [ -f "$TEMPLATE_FILE" ]; then
  # Use template
  sed -e "s|{{USER}}|$INSTALL_USER|g" \
    -e "s|{{APP_DIR}}|$APP_DIR|g" \
    -e "s|{{PYTHON_CMD}}|$PYTHON_CMD|g" \
    "$TEMPLATE_FILE" >"$SERVICE_FILE"
else
  # Generate service file directly
  cat >"$SERVICE_FILE" <<EOF
[Unit]
Description=JSON Schema Form Generator
After=network.target

[Service]
Type=simple
User=$INSTALL_USER
WorkingDirectory=$APP_DIR
ExecStart=$PYTHON_CMD $APP_DIR/app.py
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal

# Environment
Environment="PYTHONUNBUFFERED=1"

[Install]
WantedBy=multi-user.target
EOF
fi

echo -e "   Created: $SERVICE_FILE"

# Set correct permissions
chmod 644 "$SERVICE_FILE"

# Reload systemd
echo -e "${BLUE}🔄 Reloading systemd...${NC}"
systemctl daemon-reload

# Enable service
echo -e "${BLUE}✅ Enabling service...${NC}"
systemctl enable form-generator.service

# Ask if user wants to start now
echo ""
echo -e "${YELLOW}🚀 Start the service now?${NC}"
read -p "   [Y/n] " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Nn]$ ]]; then
  echo -e "${BLUE}   Starting service...${NC}"
  systemctl start form-generator.service
  sleep 2

  # Check status
  if systemctl is-active --quiet form-generator.service; then
    echo -e "${GREEN}   ✓ Service started successfully${NC}"
  else
    echo -e "${RED}   ✗ Service failed to start${NC}"
    echo -e "${YELLOW}   Check logs: sudo journalctl -u form-generator -n 50${NC}"
  fi
else
  echo -e "${YELLOW}   ⚠ Service not started. Start manually with:${NC}"
  echo -e "      sudo systemctl start form-generator"
fi

# Installation summary
echo ""
echo -e "${GREEN}============================================================${NC}"
echo -e "${GREEN}  ✓ Installation Complete!${NC}"
echo -e "${GREEN}============================================================${NC}"
echo ""
echo -e "${BLUE}📋 Service Information:${NC}"
echo "   Name:        form-generator.service"
echo "   User:        $INSTALL_USER"
echo "   Directory:   $APP_DIR"
echo "   Python:      $PYTHON_CMD"
echo ""
echo -e "${BLUE}🔧 Useful Commands:${NC}"
echo "   Start:       sudo systemctl start form-generator"
echo "   Stop:        sudo systemctl stop form-generator"
echo "   Restart:     sudo systemctl restart form-generator"
echo "   Status:      sudo systemctl status form-generator"
echo "   Logs:        sudo journalctl -u form-generator -f"
echo "   Disable:     sudo systemctl disable form-generator"
echo ""
echo -e "${BLUE}🌐 Access:${NC}"
echo "   http://localhost:8200"
echo "   http://$(hostname -I | awk '{print $1}'):8200"
echo ""
echo -e "${YELLOW}💡 Tip: Run ./install/uninstall.sh to remove the service${NC}"
echo ""
