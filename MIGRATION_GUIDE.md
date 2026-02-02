# Environment Configuration Migration Guide

## What Changed?

The environment configuration has been refactored to support multiple independent stacks on the same host. This enables running different environments (pre-production, production, development, etc.) simultaneously without conflicts.

## Before (Old Structure)

Environment files contained both shared and environment-specific settings mixed together:

```
env/
├── prod.env     # All settings for production
└── pre.env      # All settings for pre-production (duplicated many settings)
```

**Problems:**
- Configuration duplication across environment files
- Difficult to maintain consistency (e.g., SMTP settings)
- Hard to deploy multiple stacks on same host
- No clear separation between shared and environment-specific settings

## After (New Structure)

Configuration is split into shared and environment-specific files:

```
env/
├── common.env   # Shared configuration (SMTP, polling behavior, etc.)
├── prod.env     # Production-specific only (ports, paths, project name)
└── pre.env      # Pre-production-specific only (ports, paths, project name)
```

**Benefits:**
- ✅ No duplication - shared settings are in one place
- ✅ Easy to maintain - update SMTP once, applies to all environments
- ✅ Multiple stacks on same host - each has unique ports, paths, and project name
- ✅ Clear separation - instantly see what differs between environments
- ✅ Different extractor versions per environment - easy to test new versions

## Migration Steps

### 1. Update Your Deployment Commands

**Old way:**
```bash
docker compose --env-file env/prod.env up -d
```

**New way:**
```bash
docker compose --env-file env/common.env --env-file env/prod.env up -d
```

**Important:** You must now specify BOTH environment files:
1. `env/common.env` - Always first
2. `env/prod.env` (or pre.env, etc.) - Environment-specific settings

### 2. Update Your Scripts

If you have deployment scripts, update all `docker compose` commands:

```bash
# Add these flags to all docker compose commands
--env-file env/common.env --env-file env/prod.env
```

**Examples:**

```bash
# Start
docker compose --env-file env/common.env --env-file env/prod.env up -d

# Stop
docker compose --env-file env/common.env --env-file env/prod.env down

# View logs
docker compose --env-file env/common.env --env-file env/prod.env logs -f

# Rebuild
docker compose --env-file env/common.env --env-file env/prod.env up -d --build
```

### 3. Validate Configuration

Run the validation script to ensure your environments are properly configured:

```bash
./validate_multistack.sh
```

This will check:
- ✓ All required files exist
- ✓ Configurations are valid
- ✓ No conflicts between environments (ports, volumes, project names)
- ✓ Common configuration is properly shared

## What's in Each File?

### env/common.env (Shared Settings)

Contains configuration that is the same across all environments:

- **SMTP Settings:** Host, port, credentials, from address
- **Email Behavior:** Mode, default recipients
- **Watcher Settings:** Polling interval, batch quiet time, done marker
- **API Limits:** Max ZIP size

**When to edit:**
- Changing SMTP server
- Updating email recipients for all environments
- Adjusting polling behavior globally
- Changing API limits

### env/prod.env (Production-Specific)

Contains only production-specific settings:

- **COMPOSE_PROJECT_NAME:** `ias_prod` (container name prefix)
- **STACK_ROOT:** `/data/ias_prod` (data storage path)
- **Ports:**
  - SFTP_PORT: `2222`
  - SFTPGO_WEB_PORT: `8080`
  - TOOLS_WEB_PORT: `8081`

**When to edit:**
- Changing production ports
- Moving production data location
- Environment-specific overrides

### env/pre.env (Pre-Production-Specific)

Contains only pre-production-specific settings:

- **COMPOSE_PROJECT_NAME:** `ias_pre` (different from prod!)
- **STACK_ROOT:** `/data/ias_pre` (separate from prod!)
- **Ports:** Different from production to avoid conflicts
  - SFTP_PORT: `2223`
  - SFTPGO_WEB_PORT: `8083`
  - TOOLS_WEB_PORT: `8082`
- **Overrides:** May override common settings (e.g., faster `BATCH_QUIET_SECONDS` for testing)

## Running Multiple Stacks

### Deploy Both Pre and Prod

```bash
# Deploy production (ports 8080-8081)
docker compose --env-file env/common.env --env-file env/prod.env up -d

# Deploy pre-production (ports 8082-8083)
docker compose --env-file env/common.env --env-file env/pre.env up -d

# Verify both are running
docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
```

You should see containers prefixed with both `ias_prod-` and `ias_pre-`.

### Access Each Environment

**Production:**
- Web UI: http://localhost:8081/tools/lear_cable/
- API: http://localhost:8081/api/lear_cable/
- SFTP: `sftp -P 2222 lear_cable@localhost`

**Pre-production:**
- Web UI: http://localhost:8082/tools/lear_cable/
- API: http://localhost:8082/api/lear_cable/
- SFTP: `sftp -P 2223 lear_cable@localhost`

## Creating New Environments

To create a new environment (e.g., "staging" or "dev"):

1. **Copy the template:**
   ```bash
   cp .env.example env/staging.env
   ```

2. **Edit the file:**
   ```bash
   # Set unique values
   COMPOSE_PROJECT_NAME=ias_staging
   STACK_ROOT=/data/ias_staging
   SFTP_PORT=2224
   SFTPGO_WEB_PORT=8084
   TOOLS_WEB_PORT=8085
   ```

3. **Deploy:**
   ```bash
   docker compose --env-file env/common.env --env-file env/staging.env up -d
   ```

4. **Validate:**
   ```bash
   ./validate_multistack.sh
   ```

## Troubleshooting

### "Error: address already in use"

**Cause:** Port conflict between stacks or with other services.

**Solution:** Ensure each environment uses unique ports in its env file.

### Containers have unexpected names

**Cause:** Missing or incorrect COMPOSE_PROJECT_NAME.

**Solution:** Verify COMPOSE_PROJECT_NAME is set in the environment-specific file and is unique.

### Configuration doesn't take effect

**Cause:** Forgot to specify environment files in docker compose command.

**Solution:** Always use both `--env-file env/common.env --env-file env/{env}.env`

### Can't find common settings

**Cause:** Looking in environment-specific file for settings that are now in common.env.

**Solution:** Check `env/common.env` for SMTP, polling, and other shared settings.

## Quick Reference

### Deploy Commands by Environment

```bash
# Production
docker compose --env-file env/common.env --env-file env/prod.env up -d

# Pre-production
docker compose --env-file env/common.env --env-file env/pre.env up -d

# Custom environment
docker compose --env-file env/common.env --env-file env/custom.env up -d
```

### Stop Commands by Environment

```bash
# Stop production
docker compose --env-file env/common.env --env-file env/prod.env down

# Stop pre-production
docker compose --env-file env/common.env --env-file env/pre.env down
```

### View Logs by Environment

```bash
# Production logs
docker compose --env-file env/common.env --env-file env/prod.env logs -f

# Pre-production logs
docker compose --env-file env/common.env --env-file env/pre.env logs -f
```

## Additional Resources

- **Full deployment guide:** [MULTI_STACK_DEPLOYMENT.md](MULTI_STACK_DEPLOYMENT.md)
- **Usage guide:** [USAGE.md](USAGE.md)
- **Troubleshooting:** [TROUBLESHOOTING.md](TROUBLESHOOTING.md)
- **Example configuration:** [.env.example](.env.example)

## Summary of Changes

| Aspect | Before | After |
|--------|--------|-------|
| **Config files** | env/prod.env, env/pre.env | env/common.env + env/prod.env + env/pre.env |
| **Deployment command** | `docker compose --env-file env/prod.env up -d` | `docker compose --env-file env/common.env --env-file env/prod.env up -d` |
| **Multi-stack support** | ❌ Difficult, prone to conflicts | ✅ Easy, fully isolated |
| **Config duplication** | ❌ SMTP and other settings duplicated | ✅ No duplication |
| **Maintenance** | ❌ Update settings in multiple files | ✅ Update once in common.env |
| **Isolation** | ⚠️ Same project name, volume conflicts | ✅ Unique project names, separate volumes |
