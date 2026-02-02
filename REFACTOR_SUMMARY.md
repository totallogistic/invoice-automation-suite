# Environment Configuration Refactor - Summary

## Overview

This refactor reorganizes the environment configuration to enable multiple independent stacks to run on the same host, while reducing configuration duplication and improving maintainability.

## What Problem Does This Solve?

### Before
- Environment files (pre.env, prod.env) duplicated many settings (SMTP, polling, etc.)
- Difficult to run multiple environments on the same host due to port and volume conflicts
- Unclear separation between shared and environment-specific configuration
- Hard to maintain consistency across environments

### After
- Shared configuration in one place (`env/common.env`)
- Each environment has unique ports, volumes, and project names
- Clear separation of concerns
- Easy to deploy multiple stacks (pre, prod, dev, etc.) on same host
- Can run different extractor versions in different stacks

## File Structure

```
env/
├── common.env          # NEW: Shared settings (SMTP, polling, etc.)
├── prod.env            # MODIFIED: Production-specific only (ports, paths, name)
└── pre.env             # MODIFIED: Pre-production-specific only (ports, paths, name)
```

## Key Changes

### 1. New File: `env/common.env`
Contains all shared configuration:
- SMTP settings (host, port, user, password, from address)
- Email configuration (mode, default recipients)
- Watcher behavior (polling interval, batch quiet time, done marker)
- API settings (max ZIP size)

### 2. Simplified `env/prod.env`
Now contains only:
- `COMPOSE_PROJECT_NAME=ias_prod` (unique identifier)
- `STACK_ROOT=/data/ias_prod` (storage location)
- Ports (2222, 8080, 8081)
- Environment-specific overrides (if needed)

### 3. Simplified `env/pre.env`
Now contains only:
- `COMPOSE_PROJECT_NAME=ias_pre` (different from prod)
- `STACK_ROOT=/data/ias_pre` (separate storage)
- Ports (2223, 8082, 8083) - different from prod
- Environment-specific overrides (faster batch processing for testing)

### 4. Updated `compose.yaml`
Already uses `STACK_ROOT` for all volume paths - no changes needed!

### 5. Updated `.env.example`
Now a template for creating new environments with:
- Clear documentation of what's required vs optional
- Instructions on using multiple env files
- Example of unique ports and paths

## Deployment Changes

### Before
```bash
docker compose --env-file env/prod.env up -d
```

### After
```bash
docker compose --env-file env/common.env --env-file env/prod.env up -d
```

**Important:** Must specify BOTH env files in order:
1. `env/common.env` first (shared settings)
2. `env/prod.env` (or pre.env, etc.) second (environment-specific settings)

## Multi-Stack Deployment

You can now run multiple stacks on the same host:

```bash
# Deploy production
docker compose --env-file env/common.env --env-file env/prod.env up -d

# Deploy pre-production (different ports, volumes, containers)
docker compose --env-file env/common.env --env-file env/pre.env up -d
```

Each stack has:
- ✅ Unique container names (prefixed with project name)
- ✅ Unique ports (no conflicts)
- ✅ Separate volumes (no data mixing)
- ✅ Isolated networks
- ✅ Shared common configuration

## New Documentation

### MULTI_STACK_DEPLOYMENT.md
Comprehensive guide covering:
- Architecture and benefits
- Quick start for single and multi-environment deployments
- Creating new environments
- Configuration details for each variable
- Running different extractor versions
- Managing multiple stacks (logs, updates, stopping)
- Storage requirements
- Best practices
- Troubleshooting

### MIGRATION_GUIDE.md
Step-by-step migration guide covering:
- What changed and why
- Before/after comparison
- Migration steps for existing deployments
- What's in each file
- Running multiple stacks
- Creating new environments
- Troubleshooting
- Quick reference for common commands

### SECURITY.md
Security considerations covering:
- Current state (credentials in common.env)
- Recommended security improvements
- Best practices (separate credentials, Docker Secrets, etc.)
- Git history cleanup
- Recommendations by repository type
- Checklist for secure deployment
- Quick fix for current situation

### validate_multistack.sh
Validation script that checks:
- All required files exist
- Configurations are valid
- No conflicts (ports, volumes, project names)
- Common configuration is properly shared
- Provides clear success/error messages

## Updated Documentation

### README.md
- Updated quickstart with new deployment commands
- Added multi-environment deployment section
- Referenced MULTI_STACK_DEPLOYMENT.md for details

### USAGE.md
- Added note about multi-stack deployments
- Shows how to use env files in commands

### TROUBLESHOOTING.md
- Added note about multi-stack command syntax
- Referenced MULTI_STACK_DEPLOYMENT.md

### README_SECTIONS.md
- Updated all docker compose commands to use env files
- Added notes about multi-stack deployments

## Configuration Variables

### In `env/common.env` (Shared)
```bash
SMTP_HOST=...
SMTP_PORT=...
SMTP_USER=...
SMTP_PASS=...
MAIL_FROM=...
MAIL_TO=...
EMAIL_MODE=BATCH_ONLY
DONE_MARKER=_DONE
POLL_SECONDS=3
BATCH_QUIET_SECONDS=30
MAX_ZIP_MB=200
```

### In `env/prod.env` (Production-Specific)
```bash
COMPOSE_PROJECT_NAME=ias_prod
STACK_ROOT=/data/ias_prod
SFTP_PORT=2222
SFTPGO_WEB_PORT=8080
TOOLS_WEB_PORT=8081
BATCH_QUIET_SECONDS=1  # Override for faster processing
```

### In `env/pre.env` (Pre-Production-Specific)
```bash
COMPOSE_PROJECT_NAME=ias_pre
STACK_ROOT=/data/ias_pre
SFTP_PORT=2223
SFTPGO_WEB_PORT=8083
TOOLS_WEB_PORT=8082
BATCH_QUIET_SECONDS=1
WEB_SITE_DIR=./services/web/site_pre
```

## Benefits

### 1. Configuration Management
- ✅ **No duplication:** SMTP settings defined once, used everywhere
- ✅ **Easy updates:** Change SMTP in one place, applies to all environments
- ✅ **Clear ownership:** Immediately see which settings differ between environments

### 2. Multi-Stack Support
- ✅ **Run multiple stacks:** Pre, prod, dev, etc. all on same host
- ✅ **No conflicts:** Each stack has unique ports, volumes, and container names
- ✅ **Independent updates:** Update pre without affecting prod
- ✅ **Different versions:** Run different extractor versions in each stack

### 3. Maintenance
- ✅ **Single source of truth:** Common settings in one file
- ✅ **Easy to create new environments:** Copy .env.example, set unique values
- ✅ **Validation:** Script to check configurations are correct
- ✅ **Documentation:** Clear guides for deployment and migration

### 4. Flexibility
- ✅ **Override common settings:** Environment can override any common setting
- ✅ **Environment-specific behavior:** Pre can have faster batch processing
- ✅ **Custom web sites:** Different UI per environment
- ✅ **Testing:** Easy to test changes in pre before deploying to prod

## Validation

Run the validation script to verify your setup:

```bash
./validate_multistack.sh
```

Output will show:
- ✓ All required files exist
- ✓ Configurations are valid
- ✓ No conflicts between environments
- ✓ Common configuration is shared properly

## Example Use Cases

### Use Case 1: Testing New Extractor Version
```bash
# Deploy prod with stable extractor
docker compose --env-file env/common.env --env-file env/prod.env up -d

# Create dev environment with new extractor
cp .env.example env/dev.env
# Edit dev.env with unique ports/paths
# Mount different extractor version in compose override

docker compose --env-file env/common.env --env-file env/dev.env up -d
```

### Use Case 2: Separate Pre and Prod
```bash
# Run both environments simultaneously
docker compose --env-file env/common.env --env-file env/pre.env up -d
docker compose --env-file env/common.env --env-file env/prod.env up -d

# Access pre on ports 8082-8083
curl http://localhost:8082/api/lear_cable/health

# Access prod on ports 8080-8081
curl http://localhost:8081/api/lear_cable/health
```

### Use Case 3: Update SMTP for All Environments
```bash
# Edit only env/common.env
vim env/common.env  # Change SMTP_HOST

# Restart all stacks to pick up new SMTP
docker compose --env-file env/common.env --env-file env/prod.env restart
docker compose --env-file env/common.env --env-file env/pre.env restart
```

## Backwards Compatibility

⚠️ **Breaking Change:** Deployment commands must be updated to include both env files.

**Migration Required:**
- Update all deployment scripts
- Update CI/CD pipelines
- Update documentation/runbooks
- Train team on new deployment process

See [MIGRATION_GUIDE.md](MIGRATION_GUIDE.md) for detailed migration steps.

## Testing

The refactor has been validated:
- ✅ Production config loads correctly
- ✅ Pre-production config loads correctly
- ✅ Both can run simultaneously without conflicts
- ✅ All services get correct environment variables
- ✅ Volumes are properly isolated
- ✅ Ports are unique per stack
- ✅ Common configuration is shared correctly

Run `./validate_multistack.sh` to verify on your system.

## Next Steps

1. **Review the changes** in this PR
2. **Test locally** using the validation script
3. **Update deployment scripts** to use new env file syntax
4. **Consider security improvements** from SECURITY.md
5. **Train team** on new deployment process
6. **Update CI/CD** pipelines if applicable
7. **Deploy** when ready

## Files Changed

- **Created:**
  - `env/common.env` - Shared configuration
  - `MULTI_STACK_DEPLOYMENT.md` - Deployment guide
  - `MIGRATION_GUIDE.md` - Migration instructions
  - `SECURITY.md` - Security considerations
  - `validate_multistack.sh` - Validation script
  - `REFACTOR_SUMMARY.md` - This file

- **Modified:**
  - `env/prod.env` - Now contains only production-specific settings
  - `env/pre.env` - Now contains only pre-production-specific settings
  - `.env.example` - Updated template with new structure
  - `README.md` - Updated deployment instructions
  - `USAGE.md` - Added multi-stack note
  - `TROUBLESHOOTING.md` - Added multi-stack note
  - `README_SECTIONS.md` - Updated all commands

- **Unchanged:**
  - `compose.yaml` - Already supports STACK_ROOT, no changes needed
  - Services code - No application changes required

## Questions?

- **General deployment:** See [MULTI_STACK_DEPLOYMENT.md](MULTI_STACK_DEPLOYMENT.md)
- **Migration help:** See [MIGRATION_GUIDE.md](MIGRATION_GUIDE.md)
- **Security concerns:** See [SECURITY.md](SECURITY.md)
- **Validation:** Run `./validate_multistack.sh`
