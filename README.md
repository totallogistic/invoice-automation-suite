# Invoice Automation Suite - Unified Architecture

Automated invoice processing system with multi-tool support.

## Architecture

**Unified Services** (3 total):
- `unified_api` - Single API for all tools
- `unified_processor` - Single processor for all tools  
- `tools_web` - Web interface + reverse proxy

**Current Tools**:
- Lear Cable Invoice Extractor
- Import Partida Processor

## Quick Start
```bash
# 1. Configure environment
cp .env.example .env
# Edit .env with your SMTP settings

# 2. Start services
docker compose up -d

# 3. Access web interface
http://localhost:8081/
```

## Adding a New Tool

1. Create extractor:
```bash
   mkdir -p apps/my_tool/extractor
   # Write your extractor script
```

2. Add to config:
```yaml
   # config/tools.yaml
   - name: my_tool
     display_name: "My Tool"
     extractor:
       path: /apps/my_tool/extractor/process.py
     input:
       formats: [pdf]
     output:
       artifacts: [output.xlsx]
     email:
       subject_template: "[My Tool] Batch {batch_id}"
```

3. Restart:
```bash
   docker compose restart
```

## Directory Structure
```
apps/              Tool-specific extractors
config/            Tool configurations
libs/              Shared utilities
services/          Core services (API, Processor, Web)
tests/             Tests
```

## Environment Variables

Required in `.env`:
```bash
# Data
DATA_ROOT=/data

# Web
TOOLS_WEB_PORT=8081

# Email
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USER=user@example.com
SMTP_PASS=password
MAIL_FROM=noreply@example.com
MAIL_TO=recipient@example.com
```

## Development
```bash
# View logs
docker compose logs -f

# Restart after config change
docker compose restart unified_api unified_processor

# Run tests
pytest tests/
```

## Migration from Old Architecture

This repository previously used separate services per tool. 
The old code is archived in `archive/` directory.

**Benefits of unified architecture**:
- 73% fewer services
- 90% faster to add new tools
- Single codebase for infrastructure
- Easier maintenance and testing

---

For detailed documentation, see the `docs/` directory.
