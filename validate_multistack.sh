#!/bin/bash
# Validation script to verify multi-stack configuration
# This script validates that prod and pre environments are properly configured
# and can coexist on the same host

set -e

echo "==================================================="
echo "Multi-Stack Configuration Validation"
echo "==================================================="
echo ""

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print success
success() {
    echo -e "${GREEN}✓${NC} $1"
}

# Function to print error
error() {
    echo -e "${RED}✗${NC} $1"
}

# Function to print warning
warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

echo "Checking environment files..."
echo ""

# Check if common.env exists
if [ -f "env/common.env" ]; then
    success "env/common.env exists"
else
    error "env/common.env not found"
    exit 1
fi

# Check if prod.env exists
if [ -f "env/prod.env" ]; then
    success "env/prod.env exists"
else
    error "env/prod.env not found"
    exit 1
fi

# Check if pre.env exists
if [ -f "env/pre.env" ]; then
    success "env/pre.env exists"
else
    error "env/pre.env not found"
    exit 1
fi

echo ""
echo "Validating production configuration..."
echo ""

# Validate prod config
PROD_CONFIG=$(docker compose --env-file env/common.env --env-file env/prod.env config --format json 2>&1)
if [ $? -eq 0 ]; then
    success "Production configuration is valid"
    
    # Extract key values
    PROD_NAME=$(echo "$PROD_CONFIG" | jq -r '.name')
    PROD_SFTP_PORT=$(echo "$PROD_CONFIG" | jq -r '.services.sftp_lear_cable.ports[0].published')
    PROD_WEB_PORT=$(echo "$PROD_CONFIG" | jq -r '.services.tools_web.ports[0].published')
    PROD_VOLUME=$(echo "$PROD_CONFIG" | jq -r '.services.processor_lear_cable.volumes[0].source')
    
    echo "  - Project name: $PROD_NAME"
    echo "  - SFTP port: $PROD_SFTP_PORT"
    echo "  - Web port: $PROD_WEB_PORT"
    echo "  - Data volume: $PROD_VOLUME"
else
    error "Production configuration is invalid"
    exit 1
fi

echo ""
echo "Validating pre-production configuration..."
echo ""

# Validate pre config
PRE_CONFIG=$(docker compose --env-file env/common.env --env-file env/pre.env config --format json 2>&1)
if [ $? -eq 0 ]; then
    success "Pre-production configuration is valid"
    
    # Extract key values
    PRE_NAME=$(echo "$PRE_CONFIG" | jq -r '.name')
    PRE_SFTP_PORT=$(echo "$PRE_CONFIG" | jq -r '.services.sftp_lear_cable.ports[0].published')
    PRE_WEB_PORT=$(echo "$PRE_CONFIG" | jq -r '.services.tools_web.ports[0].published')
    PRE_VOLUME=$(echo "$PRE_CONFIG" | jq -r '.services.processor_lear_cable.volumes[0].source')
    
    echo "  - Project name: $PRE_NAME"
    echo "  - SFTP port: $PRE_SFTP_PORT"
    echo "  - Web port: $PRE_WEB_PORT"
    echo "  - Data volume: $PRE_VOLUME"
else
    error "Pre-production configuration is invalid"
    exit 1
fi

echo ""
echo "Checking for conflicts between environments..."
echo ""

# Check if project names are different
if [ "$PROD_NAME" != "$PRE_NAME" ]; then
    success "Project names are unique ($PROD_NAME vs $PRE_NAME)"
else
    error "Project names are the same! Containers will conflict."
    exit 1
fi

# Check if SFTP ports are different
if [ "$PROD_SFTP_PORT" != "$PRE_SFTP_PORT" ]; then
    success "SFTP ports are unique ($PROD_SFTP_PORT vs $PRE_SFTP_PORT)"
else
    error "SFTP ports are the same! Port conflict will occur."
    exit 1
fi

# Check if web ports are different
if [ "$PROD_WEB_PORT" != "$PRE_WEB_PORT" ]; then
    success "Web ports are unique ($PROD_WEB_PORT vs $PRE_WEB_PORT)"
else
    error "Web ports are the same! Port conflict will occur."
    exit 1
fi

# Check if volume paths are different
if [ "$PROD_VOLUME" != "$PRE_VOLUME" ]; then
    success "Volume paths are unique ($PROD_VOLUME vs $PRE_VOLUME)"
else
    error "Volume paths are the same! Data will be shared/corrupted."
    exit 1
fi

echo ""
echo "Checking common configuration propagation..."
echo ""

# Check if SMTP settings from common.env are in both configs
PROD_SMTP=$(echo "$PROD_CONFIG" | jq -r '.services.processor_lear_cable.environment.SMTP_HOST')
PRE_SMTP=$(echo "$PRE_CONFIG" | jq -r '.services.processor_lear_cable.environment.SMTP_HOST')

if [ "$PROD_SMTP" = "$PRE_SMTP" ] && [ "$PROD_SMTP" != "null" ]; then
    success "Common SMTP configuration is shared (SMTP_HOST: $PROD_SMTP)"
else
    warning "SMTP configuration differs or is missing between environments"
fi

echo ""
echo "==================================================="
echo -e "${GREEN}All validation checks passed!${NC}"
echo "==================================================="
echo ""
echo "Summary:"
echo "  - Both environments are properly configured"
echo "  - No conflicts detected (ports, volumes, project names)"
echo "  - Common configuration is properly shared"
echo ""
echo "You can now deploy both stacks on the same host:"
echo ""
echo "  Production:"
echo "    docker compose --env-file env/common.env --env-file env/prod.env up -d"
echo ""
echo "  Pre-production:"
echo "    docker compose --env-file env/common.env --env-file env/pre.env up -d"
echo ""
