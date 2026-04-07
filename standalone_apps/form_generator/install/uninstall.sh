#!/bin/bash
#
# Form Generator - Uninstallation Script
# Removes all form-generator systemd services
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

# Find all form-generator services
SERVICES=$(systemctl list-units --all --type=service --plain --no-legend | grep 'form-generator' | awk '{print $1}' || true)

if [ -z "$SERVICES" ]; then
  echo -e "${YELLOW}⚠ No form-generator services found. Nothing to uninstall.${NC}"
  exit 0
fi

# Show services to be removed
echo -e "${BLUE}📋 Found services:${NC}"
echo "$SERVICES" | sed 's/^/   /'
echo ""

# Confirm uninstallation
echo -e "${YELLOW}⚠️  This will remove all form-generator services.${NC}"
echo -e "${YELLOW}   The application files will NOT be deleted.${NC}"
echo ""
read -p "Continue? [y/N] " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
  echo -e "${BLUE}Cancelled.${NC}"
  exit 0
fi

echo ""

# Stop, disable and remove each service
while IFS= read -r service; do
  if [ -z "$service" ]; then
    continue
  fi

  service_name=$(basename "$service" .service)
  service_file="/etc/systemd/system/$service"

  echo -e "${BLUE}Processing: ${GREEN}$service${NC}"

  # Stop service
  if systemctl is-active --quiet "$service" 2>/dev/null; then
    systemctl stop "$service"
    echo -e "   ${GREEN}✓ Stopped${NC}"
  else
    echo -e "   ${YELLOW}⚠ Was not running${NC}"
  fi

  # Disable service
  if systemctl is-enabled --quiet "$service" 2>/dev/null; then
    systemctl disable "$service" 2>/dev/null || true
    echo -e "   ${GREEN}✓ Disabled${NC}"
  else
    echo -e "   ${YELLOW}⚠ Was not enabled${NC}"
  fi

  # Remove service file
  if [ -f "$service_file" ]; then
    rm -f "$service_file"
    echo -e "   ${GREEN}✓ Removed service file${NC}"
  fi

  echo ""
done <<<"$SERVICES"

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
echo -e "${YELLOW}💡 To reinstall: sudo ./install/install-all.sh${NC}"
echo ""

