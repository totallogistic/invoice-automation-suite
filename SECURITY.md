# Security Considerations for Environment Files

## ⚠️ Important Security Notice

The current configuration files (`env/common.env`, `env/prod.env`, `env/pre.env`) contain actual credentials and should be treated as sensitive.

## Current State

**Files committed to repository:**
- `env/common.env` - Contains SMTP credentials ⚠️
- `env/prod.env` - Contains paths and ports (less sensitive)
- `env/pre.env` - Contains paths and ports (less sensitive)

## Recommended Security Improvements

### Option 1: Template-Based Approach (Recommended for open-source)

If this repository will be public or shared:

1. **Create template files (safe to commit):**
   ```bash
   env/common.env.template
   env/prod.env.template
   env/pre.env.template
   ```

2. **Update .gitignore to exclude actual env files:**
   ```
   .env
   env/*.env
   !env/*.env.template
   ```

3. **Users copy templates on first setup:**
   ```bash
   cp env/common.env.template env/common.env
   cp env/prod.env.template env/prod.env
   # Then edit env/common.env with actual credentials
   ```

### Option 2: Keep Current Approach (For private repositories only)

If this repository will remain strictly private:

1. **Ensure repository is private** on GitHub
2. **Limit access** to only trusted team members
3. **Use repository secrets** for CI/CD instead of committing credentials
4. **Consider using environment-specific credential files** that are not committed

## Best Practices

### 1. Separate Credentials from Configuration

Create a separate credentials file:

```bash
env/credentials.env  # Not committed (in .gitignore)
```

Content:
```bash
SMTP_USER=procesos@tlseng.es
SMTP_PASS=YourSecurePassword
```

Update `env/common.env` to not include credentials:
```bash
# SMTP Configuration
SMTP_HOST=tlseng-es.correoseguro.dinaserver.com
SMTP_PORT=465
# SMTP_USER and SMTP_PASS are in env/credentials.env
```

Deploy with three env files:
```bash
docker compose \
  --env-file env/credentials.env \
  --env-file env/common.env \
  --env-file env/prod.env \
  up -d
```

### 2. Use Docker Secrets (For Production)

For production deployments, use Docker Secrets:

```yaml
secrets:
  smtp_password:
    external: true

services:
  processor_lear_cable:
    secrets:
      - smtp_password
    environment:
      - SMTP_PASS_FILE=/run/secrets/smtp_password
```

Create the secret:
```bash
echo "YourSecurePassword" | docker secret create smtp_password -
```

### 3. Use Environment Variables

For quick testing, use environment variables:

```bash
export SMTP_PASS="YourSecurePassword"
docker compose --env-file env/common.env --env-file env/prod.env up -d
```

Ensure `env/common.env` references the variable:
```bash
SMTP_PASS=${SMTP_PASS}
```

### 4. Credential Management Tools

Consider using:
- **HashiCorp Vault** - Enterprise credential management
- **AWS Secrets Manager** - For AWS deployments
- **Azure Key Vault** - For Azure deployments
- **Google Secret Manager** - For GCP deployments

## Current Credentials in Repository

⚠️ **Action Required:**

The following sensitive information is currently in `env/common.env`:

1. **SMTP Credentials:** Username and password are in plaintext
2. **Email addresses:** Various recipients

**Recommended actions:**

1. **Rotate the SMTP password** - Current password should be changed
2. **Remove credentials from env files** - Use one of the approaches above
3. **Review repository access** - Ensure only authorized users have access
4. **Audit who has cloned the repository** - Anyone with access has seen the credentials

## Git History Cleanup

If credentials were committed by mistake, they remain in git history. To remove:

```bash
# WARNING: This rewrites history and requires force push
# Coordinate with team before running

# Remove env/common.env from history
git filter-branch --force --index-filter \
  "git rm --cached --ignore-unmatch env/common.env" \
  --prune-empty --tag-name-filter cat -- --all

# Force push (after team coordination)
git push origin --force --all
```

**Better approach:** Rotate the exposed credentials instead of rewriting history.

## Recommendations by Repository Type

### Private Repository (Current State)
✅ Can keep credentials in env files (current approach)
⚠️ Ensure repository remains private
⚠️ Limit access to trusted team members only
✅ Consider rotating credentials periodically

### Public Repository
❌ Never commit credentials
✅ Use `.env.template` files
✅ Document what credentials are needed
✅ Provide example values (not real ones)

### Enterprise Deployment
✅ Use Docker Secrets or external credential management
✅ Implement credential rotation
✅ Use separate credential files not in git
✅ Audit access to credentials

## Checklist for Secure Deployment

- [ ] Review who has access to this repository
- [ ] Decide on credential management approach (see options above)
- [ ] Rotate any exposed credentials
- [ ] Update .gitignore if moving to template-based approach
- [ ] Document credential setup process for new team members
- [ ] Set up monitoring for unauthorized access attempts
- [ ] Implement credential rotation schedule (e.g., every 90 days)

## Quick Fix for Current Situation

If you need to secure credentials immediately:

1. **Move credentials to a separate file:**
   ```bash
   # Create credentials.env (add to .gitignore)
   cat > env/credentials.env <<EOF
   SMTP_USER=your-smtp-user@example.com
   SMTP_PASS=YourSecurePassword123
   EOF
   ```

2. **Update .gitignore:**
   ```bash
   echo "env/credentials.env" >> .gitignore
   ```

3. **Remove credentials from env/common.env:**
   ```bash
   # In env/common.env, replace:
   SMTP_USER=your-smtp-user@example.com
   SMTP_PASS=YourSecurePassword123
   
   # With comments:
   # SMTP_USER and SMTP_PASS are in env/credentials.env (not committed)
   ```

4. **Update deployment commands:**
   ```bash
   docker compose \
     --env-file env/credentials.env \
     --env-file env/common.env \
     --env-file env/prod.env \
     up -d
   ```

5. **Commit the changes:**
   ```bash
   git add .gitignore env/common.env
   git commit -m "Remove credentials from common.env, move to separate file"
   ```

6. **Rotate the SMTP password** (it's in git history now)

## Questions?

If you have questions about securing your deployment:
1. Review the [Docker Secrets documentation](https://docs.docker.com/engine/swarm/secrets/)
2. Consult your security team
3. Consider a security audit of your deployment
