#!/bin/bash
#
# Form Generator - Uninstallation Script
# Removes the systemd service
#
# Usage:
#   cd standalone_apps/form_generator
#   sudo ./install/uninstall.sh
#

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

echo -e "${BLUE}"
echo "============================================================"
echo "  🗑️  Form Generator - Uninstallation Script"
echo "============================================================"
echo -e "${NC}"

# Check if running as root
if [ "$EUID" -ne 0 ]; then 
    echo -e "${RED}❌ Please run as root (use sudo)${NC}"
    exit 1
fi

SERVICE_FILE="/etc/systemd/system/form-generator.service"

# Check if service exists
if [ ! -f "$SERVICE_FILE" ]; then
    echo -e "${YELLOW}⚠ Service not found. Nothing to uninstall.${NC}"
    exit 0
fi

# Confirm uninstallation
echo -e "${YELLOW}⚠️  This will remove the form-generator service.${NC}"
echo -e "${YELLOW}   The application files will NOT be deleted.${NC}"
echo ""
read -p "Continue? [y/N] " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo -e "${BLUE}Cancelled.${NC}"
    exit 0
fi

echo ""
echo -e "${BLUE}🛑 Stopping service...${NC}"
if systemctl is-active --quiet form-generator.service; then
    systemctl stop form-generator.service
    echo -e "${GREEN}   ✓ Service stopped${NC}"
else
    echo -e "${YELLOW}   ⚠ Service was not running${NC}"
fi

echo -e "${BLUE}🔓 Disabling service...${NC}"
if systemctl is-enabled --quiet form-generator.service 2>/dev/null; then
    systemctl disable form-generator.service
    echo -e "${GREEN}   ✓ Service disabled${NC}"
else
    echo -e "${YELLOW}   ⚠ Service was not enabled${NC}"
fi

echo -e "${BLUE}🗑️  Removing service file...${NC}"
rm -f "$SERVICE_FILE"
echo -e "${GREEN}   ✓ Removed $SERVICE_FILE${NC}"

echo -e "${BLUE}🔄 Reloading systemd...${NC}"
systemctl daemon-reload
systemctl reset-failed 2>/dev/null || true

echo ""
echo -e "${GREEN}============================================================${NC}"
echo -e "${GREEN}  ✓ Uninstallation Complete!${NC}"
echo -e "${GREEN}============================================================${NC}"
echo ""
echo -e "${BLUE}📝 Note:${NC}"
echo "   Application files remain in:"
echo "   $(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
echo ""
echo -e "${YELLOW}💡 To reinstall: sudo ./install/install.sh${NC}"
echo ""