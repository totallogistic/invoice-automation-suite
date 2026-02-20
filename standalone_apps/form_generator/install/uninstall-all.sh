#!/bin/bash
# Uninstall all form-generator services

set -e

if [ "$EUID" -ne 0 ]; then
  echo "Please run as root"
  exit 1
fi

echo "Uninstalling all Form Generator services..."
echo ""

for service in /etc/systemd/system/form-generator-*.service; do
  if [ -f "$service" ]; then
    service_name=$(basename "$service")
    echo "Removing: $service_name"

    systemctl stop "$service_name" 2>/dev/null || true
    systemctl disable "$service_name" 2>/dev/null || true
    rm -f "$service"
  fi
done

systemctl daemon-reload
systemctl reset-failed 2>/dev/null || true

echo ""
echo "✓ All Form Generator services removed"
