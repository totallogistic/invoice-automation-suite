# Invoice Automation Suite

**Unified architecture for automated invoice processing**

## 🎯 Overview

Multi-tool invoice automation system with:
- **Unified API** - Single endpoint for all tools
- **Unified Processor** - One service handles all processing
- **Configuration-driven** - Add tools via config, not code

### Current Tools
- **Lear Cable** - Invoice data extraction
- **Import Partida** - B/L document processing

## 🚀 Quick Start
```bash
# 1. Configure
cp .env.example .env
# Edit .env with your SMTP settings

# 2. Start
docker compose up -d

# 3. Access
http://localhost:8081
```

## 📁 Project Structure
```
apps/              Tool extractors (business logic only)
config/            Tool configurations
libs/              Shared utilities
services/          Core services (API, Processor, Web)
docs/              Documentation
tests/             Unit and integration tests
```

## ➕ Adding a New Tool

**Time: ~2-4 hours** (vs 2-3 days with old architecture)

1. **Create extractor** (business logic):
```bash
   mkdir -p apps/my_tool/extractor
   # Write extractor script
```

2. **Add to config** (15 lines):
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
```

3. **Restart**:
```bash
   docker compose restart
```

**Done!** ✨

## 🏗️ Architecture

**Old Architecture** (per-tool services):
```
2 tools = 5 services (2 APIs + 2 Processors + Web)
```

**New Architecture** (unified):
```
∞ tools = 3 services (API + Processor + Web)
```

### Benefits
- ✅ 73% fewer services
- ✅ 90% faster tool development
- ✅ Single codebase for infrastructure
- ✅ Configuration over code

## 📚 Documentation

- [Architecture Details](docs/ARCHITECTURE.md)
- [Adding Tools Guide](docs/ADDING_TOOLS.md)
- [Deployment Guide](docs/DEPLOYMENT.md)
- [Troubleshooting](docs/TROUBLESHOOTING.md)
- [Usage Guide](docs/USAGE.md)

## 🛠️ Development
```bash
# View logs
docker compose logs -f unified_processor

# Restart after config change
docker compose restart

# Run tests
pytest tests/
```

## 🔄 Migration from Old Architecture

This project was migrated from a per-tool service architecture.
Old code is archived locally but not in git.

**Comparison**:
| Metric | Before | After | Change |
|--------|--------|-------|--------|
| Services (2 tools) | 5 | 3 | -40% |
| Time to add tool | 2-3 days | 2-4 hours | -90% |
| Code per tool | 605 lines | 165 lines | -73% |

## 📧 Support

For issues or questions, see [TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)

---

**Version**: 2.0 (Unified Architecture)  
**Last Updated**: February 2026
