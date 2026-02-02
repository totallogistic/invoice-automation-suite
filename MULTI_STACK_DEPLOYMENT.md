# Multi-Stack Deployment Guide

This guide explains how to deploy multiple independent stacks of the Invoice Automation Suite on the same host.

> **⚠️ Security Note:** The current `env/common.env` contains credentials. For production use, consider securing credentials using Docker Secrets or separate credential files. See [SECURITY.md](SECURITY.md) for detailed recommendations.

## Architecture Overview

The refactored configuration system separates concerns into two types of files:

1. **Common Configuration** (`env/common.env`): Shared settings across all environments (SMTP, polling behavior, etc.)
2. **Environment-Specific Configuration** (`env/{env}.env`): Unique settings per environment (ports, paths, project name)

This separation allows you to:
- Run multiple independent stacks (pre, prod, dev, etc.) on the same host
- Each stack has isolated volumes, containers, and network ports
- Share common configuration while customizing per environment
- Deploy different versions of the extractor or services in each stack

## File Structure

```
invoice-automation-suite/
├── env/
│   ├── common.env      # Shared configuration (SMTP, behavior, etc.)
│   ├── prod.env        # Production-specific (ports, paths, project name)
│   ├── pre.env         # Pre-production-specific
│   └── dev.env         # Development-specific (example)
├── compose.yaml        # Docker Compose configuration (uses env vars)
└── .env.example        # Template for new environments
```

## Quick Start

### Single Environment Deployment

1. **Deploy production:**
   ```bash
   docker compose --env-file env/common.env --env-file env/prod.env up -d
   ```

2. **Verify:**
   ```bash
   docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
   curl http://localhost:8081/api/lear_cable/health
   ```

### Multi-Environment Deployment (Same Host)

1. **Deploy production stack:**
   ```bash
   docker compose --env-file env/common.env --env-file env/prod.env up -d
   ```

2. **Deploy pre-production stack (in parallel):**
   ```bash
   docker compose --env-file env/common.env --env-file env/pre.env up -d
   ```

3. **Verify both stacks:**
   ```bash
   # Check containers (should see both ias_prod-* and ias_pre-* containers)
   docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
   
   # Check prod endpoints
   curl http://localhost:8081/api/lear_cable/health
   
   # Check pre endpoints
   curl http://localhost:8082/api/lear_cable/health
   ```

## Creating a New Environment

To create a new environment (e.g., "staging"):

1. **Copy the example:**
   ```bash
   cp .env.example env/staging.env
   ```

2. **Edit `env/staging.env`:**
   ```bash
   # Set unique project name
   COMPOSE_PROJECT_NAME=ias_staging
   
   # Set unique storage path
   STACK_ROOT=/data/ias_staging
   
   # Set unique ports (must not conflict with other stacks)
   SFTP_PORT=2225
   SFTPGO_WEB_PORT=8085
   TOOLS_WEB_PORT=8086
   ```

3. **Deploy:**
   ```bash
   docker compose --env-file env/common.env --env-file env/staging.env up -d
   ```

## Configuration Details

### COMPOSE_PROJECT_NAME
- **Purpose:** Unique identifier for the stack
- **Effect:** Prefixes all container names, network names, and volumes
- **Example:** `ias_prod` → containers named `ias_prod-api_lear_cable-1`, etc.
- **Requirement:** Must be unique per stack on the same host

### STACK_ROOT
- **Purpose:** Base directory for all persistent data
- **Structure:**
  ```
  /data/ias_prod/          # STACK_ROOT
  ├── data/                # Application data (inbox, out, processed, etc.)
  ├── sftpgo/              # SFTPGo file storage
  └── sftpgo_var/          # SFTPGo database and config
  ```
- **Isolation:** Each stack has completely separate data
- **Requirement:** Must be unique per stack to avoid data conflicts

### Ports
Each stack must use unique ports:

| Service | Prod (env/prod.env) | Pre (env/pre.env) | Your Stack |
|---------|---------------------|-------------------|------------|
| SFTP | 2222 | 2223 | 2224+ |
| SFTPGo Admin | 8080 | 8083 | 8084+ |
| Web/API | 8081 | 8082 | 8085+ |

## Common Configuration (`env/common.env`)

Settings shared across all environments:
- **SMTP Configuration:** Host, port, credentials, from address
- **Email Behavior:** Mode, default recipients
- **Watcher Behavior:** Polling interval, batch quiet time, done marker
- **API Limits:** Max ZIP size

### Override Common Settings
If an environment needs different settings (e.g., different SMTP server for dev), add the variable to the environment-specific file:

```bash
# In env/dev.env
SMTP_HOST=localhost
SMTP_PORT=1025
MAIL_TO=dev@example.com
```

Environment-specific values override common values.

## Deploying Different Extractor Versions

To run different versions of the extractor in different stacks:

1. **Create version-specific extractor directories:**
   ```bash
   apps/lear_cable/extractor_v1/
   apps/lear_cable/extractor_v2/
   ```

2. **Create custom compose files:**
   ```bash
   # compose.prod.yaml - uses extractor_v1
   # compose.pre.yaml - uses extractor_v2
   ```

3. **Or use environment-specific volume mounts:**
   ```yaml
   # In environment-specific override
   volumes:
     - ${EXTRACTOR_PATH:-./apps/lear_cable/extractor}:/app/extractor:ro
   ```

## Managing Multiple Stacks

### List All Stacks
```bash
docker ps --format "table {{.Names}}\t{{.Image}}\t{{.Status}}\t{{.Ports}}" | grep ias_
```

### Stop Specific Stack
```bash
# Stop prod
docker compose --env-file env/common.env --env-file env/prod.env down

# Stop pre
docker compose --env-file env/common.env --env-file env/pre.env down
```

### Logs for Specific Stack
```bash
# Prod logs
docker compose --env-file env/common.env --env-file env/prod.env logs -f processor_lear_cable

# Pre logs
docker compose --env-file env/common.env --env-file env/pre.env logs -f processor_lear_cable
```

### Update Specific Stack
```bash
# Update prod
docker compose --env-file env/common.env --env-file env/prod.env up -d --build

# Update pre
docker compose --env-file env/common.env --env-file env/pre.env up -d --build
```

## Storage Requirements

Each stack requires disk space for:
- Application data: Varies by usage (invoices, processed files)
- SFTPGo storage: Varies by SFTP usage
- SFTPGo database: ~50MB
- Docker images: ~500MB (shared across stacks)

**Example for 2 stacks:**
```
/data/
├── ias_prod/          # Production stack
│   ├── data/          # ~5GB (example)
│   ├── sftpgo/        # ~2GB (example)
│   └── sftpgo_var/    # ~50MB
└── ias_pre/           # Pre-production stack
    ├── data/          # ~1GB (example)
    ├── sftpgo/        # ~500MB (example)
    └── sftpgo_var/    # ~50MB
```

## Network Isolation

Stacks are isolated by Docker Compose project name:
- Each stack gets its own Docker network
- Containers in one stack cannot directly communicate with containers in another stack
- If cross-stack communication is needed, use host networking or external networks

## Best Practices

1. **Always use both env files:**
   ```bash
   docker compose --env-file env/common.env --env-file env/{env}.env [command]
   ```

2. **Document custom settings:**
   Add comments in environment-specific files explaining why settings differ

3. **Port allocation:**
   Reserve port ranges for each environment (e.g., 8080-8089 for prod, 8090-8099 for pre)

4. **Backup strategy:**
   Back up each STACK_ROOT separately:
   ```bash
   tar -czf ias_prod_backup_$(date +%Y%m%d).tar.gz /data/ias_prod
   tar -czf ias_pre_backup_$(date +%Y%m%d).tar.gz /data/ias_pre
   ```

5. **Monitoring:**
   Set up monitoring per stack to track resource usage and issues independently

6. **Testing:**
   Always test deployment changes in pre-production before applying to production

## Troubleshooting

### Port Conflicts
**Symptom:** `Error: bind: address already in use`

**Solution:** Ensure each stack uses unique ports in its env file

### Container Name Conflicts
**Symptom:** `Error: container name already in use`

**Solution:** Ensure each stack has a unique COMPOSE_PROJECT_NAME

### Volume Data Mixing
**Symptom:** Data from different stacks appears mixed

**Solution:** Verify STACK_ROOT is unique per stack and properly set in env files

### Wrong Stack Being Updated
**Symptom:** Commands affect the wrong stack

**Solution:** Always specify both env files in commands:
```bash
docker compose --env-file env/common.env --env-file env/prod.env [command]
```

## Migration from Old Setup

If migrating from the old configuration:

1. **Backup existing data:**
   ```bash
   tar -czf backup_before_migration.tar.gz /data/ias_prod /data/ias_pre
   ```

2. **Update deployment scripts:**
   Change from:
   ```bash
   docker compose --env-file env/prod.env up -d
   ```
   To:
   ```bash
   docker compose --env-file env/common.env --env-file env/prod.env up -d
   ```

3. **Verify configuration:**
   ```bash
   # Check what environment variables will be used
   docker compose --env-file env/common.env --env-file env/prod.env config | grep -E "SMTP|STACK|PORT"
   ```

4. **Test before full deployment:**
   Test with one stack first, verify it works, then deploy others

## Summary

The new configuration structure provides:
- ✅ Clean separation between shared and environment-specific settings
- ✅ Easy deployment of multiple independent stacks on the same host
- ✅ Ability to run different versions of services in different stacks
- ✅ Reduced duplication and easier maintenance
- ✅ Clear documentation of what each environment customizes
